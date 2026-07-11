import { CoreMessage, streamText, tool } from 'ai';
import { createOpenAI } from '@ai-sdk/openai';
import { z } from 'zod';

// OpenRouter позволяет использовать модели бесплатно и без региональных ограничений
const openrouter = createOpenAI({
  baseURL: 'https://openrouter.ai/api/v1',
  apiKey: process.env.OPENROUTER_API_KEY,
});
// Но в Vercel AI SDK (начиная с 3.1) встроен очень мощный механизм Tool Calling (вызова функций),
// который отлично справляется с задачей маршрутизации и Human-in-the-loop (за счет возврата инструмента клиенту).

// Здесь мы имитируем вызовы к нашему MCP серверу 1С.
// В реальности вы будете использовать MCP SDK Client для передачи вызовов:
// const mcpClient = new Client({ transport: new StdioClientTransport(...) });

export async function POST(req: Request) {
  try {
    const { messages }: { messages: CoreMessage[] } = await req.json();

    const result = await streamText({
      model: openrouter('openrouter/free'), // Автоматически подбирает бесплатную модель с поддержкой Tool Calling
      messages,
      maxSteps: 5, // Разрешаем модели делать до 5 шагов (вызовов функций) подряд без ответа пользователя
      system: `Ты — ИИ-ассистент, встроенный в Outlook. Твоя цель — помогать менеджеру обрабатывать письма и работать со сделками в 1С CRM.

Правила:
1. Сначала проанализируй письмо. Письмо относится к сделкам, если в нем обсуждается коммерческое сотрудничество, услуги, разработка, проекты, закупки или клиенты (даже если пока нет суммы или деталей). Если это явный спам (скидки, рассылки) или личное (пошли на обед) — так и скажи.
2. Если письмо относится к сделке, попробуй найти ее с помощью точного поиска (search_deal_exact) по ID или имени клиента.
3. Если точный поиск не дал результатов (или ты не уверен в названии), ОБЯЗАТЕЛЬНО используй семантический поиск (search_similar_deals_rag), передав туда весь смысл письма.
4. ВАЖНО: Если похожие сделки найдены, ТЫ ОБЯЗАН остановиться и вывести сообщение: "Я не нашел точного совпадения, но нашел похожие сделки: [Перечисли ВСЕ сделки, которые вернул инструмент, указывая ID сделки и название клиента]". Не используй ссылки, пиши просто ID. НЕ задавай вопрос "Привязать или создать новую?", так как у пользователя автоматически появятся кнопки для выбора. НЕ создавай сделку без нажатия кнопки пользователем!
5. Если пользователь просит создать сделку (или если похожих вообще нет), проверь, хватает ли данных. ВНИМАНИЕ: Название клиента и описание ОБЯЗАТЕЛЬНО извлеки и придумай сам из текста письма или темы! НЕ переспрашивай их у пользователя, если можешь понять смысл из контекста. Спрашивай пользователя ТОЛЬКО о тех данных, которых абсолютно нигде нет (например, сумма). Если сумма неизвестна, предложи пользователю указать 0 или пропустить этот шаг (ты сам поставишь 0). Выводи текст чистым, без странных символов или оборванных предложений.
6. Только когда у тебя есть все данные и прямое согласие, вызывай create_deal. Если пользователь просит привязать письмо к существующей сделке, вызывай инструмент link_deal, обязательно передавая ID сделки. Если пользователь просит отменить или удалить сделку, вызывай delete_deal с указанием ID. Если пользователь просит обновить сумму или изменить стоимость существующей сделки, вызывай update_deal_amount с указанием ID и новой суммы.
7. ВАЖНО: ПОСЛЕ того как инструмент create_deal, link_deal, delete_deal или update_deal_amount успешно выполнится, ТЫ ОБЯЗАН написать финальное текстовое сообщение пользователю (например: "Сделка успешно создана/привязана/удалена/обновлена..."). Выведи детали красиво, но НИКОГДА не выводи ссылки (никаких http, https, markdown-ссылок и текстов "Ссылка на сделку"). Оставь только ID, Название клиента, Описание и Сумму. Не молчи после вызова инструмента!`,
      tools: {
        search_deal_exact: tool({
          description: 'Поиск сделки в 1С по точному совпадению (номер, ИНН, название клиента)',
          parameters: z.object({
            query: z.string().describe('Поисковой запрос'),
          }),
          execute: async ({ query }) => {
            // Имитация вызова MCP-сервера -> 1C
            console.log(`[MCP Call] search_deal_exact: ${query}`);
            if (query.includes('1001') || query.toLowerCase().includes('ромашка')) {
              return { success: true, deal: { id: "1001", client: "ООО Ромашка", amount: 150000, description: "Закупка серверов" } };
            }
            return { success: false, message: "Сделка не найдена точным поиском" };
          },
        }),
        search_similar_deals_rag: tool({
          description: 'Семантический поиск похожих сделок (RAG). Используйте, если точный поиск ничего не нашел.',
          parameters: z.object({
            semantic_query: z.string().describe('Текст письма или запрос для поиска по смыслу'),
          }),
          execute: async ({ semantic_query }) => {
            // Имитация вызова MCP-сервера -> Vector DB
            console.log(`[MCP Call] search_similar_deals_rag: ${semantic_query}`);
            return { 
              success: true, 
              message: "Найдены похожие сделки", 
              deals: [
                { id: "1003", client: "ЗАО Вектор", description: "Внедрение ПО и обучение" },
                { id: "1004", client: "Global Tech", description: "Разработка портала" }
              ] 
            };
          },
        }),
        create_deal: tool({
          description: 'Создание новой сделки в 1С',
          parameters: z.object({
            client: z.string().describe('Название клиента'),
            amount: z.number().describe('Сумма сделки'),
            description: z.string().describe('Описание сделки'),
          }),
          execute: async ({ client, amount, description }) => {
            // Имитация вызова MCP-сервера -> 1C
            console.log(`[MCP Call] create_deal: ${client}, ${amount}, ${description}`);
            const id = Math.floor(Math.random() * 10000).toString();
            return { success: true, deal: { id, client, amount, description } };
          },
        }),
        link_deal: tool({
          description: 'Привязать текущее письмо к существующей сделке в 1С',
          parameters: z.object({
            deal_id: z.string().describe('Уникальный ID сделки в 1С (например, 1003)'),
          }),
          execute: async ({ deal_id }) => {
            console.log(`[MCP Call] link_deal: ID ${deal_id}`);
            return { success: true, message: `Письмо успешно привязано к сделке ID: ${deal_id}` };
          },
        }),
        delete_deal: tool({
          description: 'Удаление (отмена) сделки из 1С',
          parameters: z.object({
            deal_id: z.string().describe('Уникальный ID сделки в 1С'),
          }),
          execute: async ({ deal_id }) => {
            // Имитация вызова MCP-сервера -> 1C
            console.log(`[MCP Call] delete_deal: ID ${deal_id}`);
            return { success: true, message: `Сделка с ID ${deal_id} успешно отменена/удалена.` };
          },
        }),
        update_deal_amount: tool({
          description: 'Обновление суммы существующей сделки в 1С',
          parameters: z.object({
            deal_id: z.string().describe('Уникальный ID сделки в 1С'),
            amount: z.number().describe('Новая сумма сделки'),
          }),
          execute: async ({ deal_id, amount }) => {
            // Имитация вызова MCP-сервера -> 1C
            console.log(`[MCP Call] update_deal_amount: ID ${deal_id}, amount: ${amount}`);
            return { success: true, deal: { id: deal_id, amount }, message: `Сумма сделки ID ${deal_id} успешно обновлена на ${amount} руб.` };
          },
        }),
      },
    });

    return result.toDataStreamResponse();
  } catch (error) {
    console.error("Chat API Error:", error);
    return new Response(JSON.stringify({ error: "Internal Server Error", details: error }), { status: 500 });
  }
}
