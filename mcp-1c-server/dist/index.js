import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema, } from "@modelcontextprotocol/sdk/types.js";
import { MemoryVectorStore } from "langchain/vectorstores/memory";
import { HuggingFaceInferenceEmbeddings } from "@langchain/community/embeddings/hf";
import { Document } from "@langchain/core/documents";
// NOTE: Убедитесь, что у вас запущена локально Ollama
const server = new Server({
    name: "1c-crm-mcp-server",
    version: "1.0.0",
}, {
    capabilities: {
        tools: {},
    },
});
// Моковые данные 1С
const mockDeals = [
    { id: "1001", client: "ООО Ромашка", amount: 150000, description: "Закупка серверов и сетевого оборудования" },
    { id: "1002", client: "ИП Иванов", amount: 50000, description: "Настройка сети и маршрутизаторов" },
    { id: "1003", client: "ЗАО Вектор", amount: 1200000, description: "Внедрение 1С ERP и обучение сотрудников" },
    { id: "1004", client: "Global Tech", amount: 300000, description: "Разработка корпоративного веб-сайта портала" }
];
// Инициализация RAG (векторной БД в памяти)
let vectorStore;
async function initVectorStore() {
    try {
        // Облачная модель векторизации от HuggingFace (без региональных блокировок)
        const embeddings = new HuggingFaceInferenceEmbeddings({
            apiKey: process.env.HF_TOKEN,
        });
        const docs = mockDeals.map(deal => new Document({
            pageContent: `Клиент: ${deal.client}. Описание: ${deal.description}. Сумма: ${deal.amount}`,
            metadata: { id: deal.id, client: deal.client, amount: deal.amount }
        }));
        vectorStore = await MemoryVectorStore.fromDocuments(docs, embeddings);
        console.error("Vector store initialized successfully.");
    }
    catch (err) {
        console.error("Failed to initialize Vector store. Did you set OPENAI_API_KEY?", err);
    }
}
server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: [
            {
                name: "search_deal_exact",
                description: "Поиск сделки в 1С по точному совпадению (ID, имени клиента или описанию)",
                inputSchema: {
                    type: "object",
                    properties: {
                        query: {
                            type: "string",
                            description: "Поисковой запрос (номер сделки, ИНН, название клиента)",
                        },
                    },
                    required: ["query"],
                },
            },
            {
                name: "search_similar_deals_rag",
                description: "Семантический поиск похожих сделок (RAG). Используйте, если точный поиск ничего не нашел.",
                inputSchema: {
                    type: "object",
                    properties: {
                        semantic_query: {
                            type: "string",
                            description: "Текст письма или вольный запрос для поиска по смыслу",
                        },
                    },
                    required: ["semantic_query"],
                },
            },
            {
                name: "create_deal",
                description: "Создание новой сделки в 1С",
                inputSchema: {
                    type: "object",
                    properties: {
                        client: {
                            type: "string",
                            description: "Название клиента",
                        },
                        amount: {
                            type: "number",
                            description: "Сумма сделки",
                        },
                        description: {
                            type: "string",
                            description: "Описание сделки",
                        },
                    },
                    required: ["client", "description"],
                },
            },
        ],
    };
});
server.setRequestHandler(CallToolRequestSchema, async (request) => {
    if (request.params.name === "search_deal_exact") {
        const query = request.params.arguments?.query;
        if (!query)
            throw new Error("Missing query parameter");
        const lowerQuery = query.toLowerCase();
        const deal = mockDeals.find(d => d.id.includes(lowerQuery) || d.client.toLowerCase().includes(lowerQuery));
        if (deal) {
            return {
                content: [{ type: "text", text: JSON.stringify({ success: true, deal, link: `https://1c.example.com/deals/${deal.id}` }) }]
            };
        }
        else {
            return {
                content: [{ type: "text", text: JSON.stringify({ success: false, message: "Сделка не найдена точным поиском" }) }]
            };
        }
    }
    if (request.params.name === "search_similar_deals_rag") {
        const semanticQuery = request.params.arguments?.semantic_query;
        if (!semanticQuery)
            throw new Error("Missing semantic_query parameter");
        if (!vectorStore) {
            return {
                content: [{ type: "text", text: JSON.stringify({ success: false, message: "RAG система не инициализирована (возможно не задан API ключ)." }) }]
            };
        }
        // Ищем топ 2 похожие сделки
        const results = await vectorStore.similaritySearch(semanticQuery, 2);
        if (results.length > 0) {
            const deals = results.map((r) => ({
                id: r.metadata.id,
                client: r.metadata.client,
                description: r.pageContent,
                link: `https://1c.example.com/deals/${r.metadata.id}`
            }));
            return {
                content: [{ type: "text", text: JSON.stringify({ success: true, message: "Найдены похожие сделки", deals }) }]
            };
        }
        else {
            return {
                content: [{ type: "text", text: JSON.stringify({ success: false, message: "Похожих сделок не найдено" }) }]
            };
        }
    }
    if (request.params.name === "create_deal") {
        const client = request.params.arguments?.client;
        const amount = request.params.arguments?.amount || 0;
        const description = request.params.arguments?.description;
        const newDeal = {
            id: Math.floor(Math.random() * 10000).toString(),
            client,
            amount,
            description
        };
        mockDeals.push(newDeal);
        // Добавляем новую сделку в RAG
        if (vectorStore) {
            const newDoc = new Document({
                pageContent: `Клиент: ${client}. Описание: ${description}. Сумма: ${amount}`,
                metadata: { id: newDeal.id, client, amount }
            });
            await vectorStore.addDocuments([newDoc]);
        }
        return {
            content: [{ type: "text", text: JSON.stringify({ success: true, deal: newDeal, link: `https://1c.example.com/deals/${newDeal.id}` }) }]
        };
    }
    throw new Error(`Tool not found: ${request.params.name}`);
});
async function run() {
    await initVectorStore();
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("1C MCP Server running on stdio");
}
run().catch(console.error);
