# 📧 AI CRM Agent for Outlook

Интеллектуальная надстройка (Add-in) для Microsoft Outlook, которая позволяет менеджерам обрабатывать входящие письма и управлять сделками в 1С CRM прямо из интерфейса почтового клиента с помощью ИИ-ассистента.

## 🌟 Обзор проекта

AI CRM Agent встраивается в боковую панель Outlook. При открытии письма плагин извлекает его контекст (отправителя, тему, текст) и передает ИИ. Ассистент анализирует текст, ищет связанные сделки в 1С (используя точный и семантический RAG-поиск) и предлагает пользователю интерактивные кнопки для привязки письма к существующей сделке или создания новой. 

Проект построен на принципах **Human-in-the-loop**: ИИ подготавливает данные, но любые изменения в CRM происходят только после явного подтверждения пользователя.

## ✨ Ключевые возможности

* **Умная классификация:** ИИ автоматически определяет, относится ли письмо к коммерческой деятельности или является спамом/личной перепиской.
* **Двухуровневый поиск сделок:**
  * *Точный поиск:* Поиск по ID или точному названию компании.
  * *Семантический поиск (RAG):* Поиск по смыслу письма с использованием векторной базы данных и эмбеддингов HuggingFace, если точный поиск не дал результатов.
* **Интерактивный UI (Fluent UI):** Интерфейс выглядит как нативная часть экосистемы Microsoft. Динамические кнопки быстрых действий (создать, привязать) сохраняют историю выбора (зеленый/серый контур).
* **Полный цикл управления сделкой (Tool Calling):**
  * Создание новой сделки (с авто-извлечением данных из текста).
  * Привязка письма к существующей сделке.
  * Обновление суммы сделки.
  * Удаление (отмена) сделки.
* **Глубокая интеграция с Outlook:** Поддержка закрепления панели (Pinning) — плагин остается открытым при переключении между письмами.

## 🏗 Архитектура и Технологии

Проект разделен на фронтенд-оркестратор и изолированный микросервис интеграции, общающиеся по протоколу **MCP (Model Context Protocol)**.

### Технологический стек
* **Frontend / UI:** Next.js (App Router), React, Fluent UI React v9, Office.js.
* **AI Orchestration:** Vercel AI SDK (`useChat`, `streamText`, `tool`).
* **LLM Provider:** OpenRouter (бесплатные модели с поддержкой Tool Calling, обход региональных ограничений).
* **Backend Integration (MCP Server):** Node.js, TypeScript, Model Context Protocol SDK.
* **RAG (Retrieval-Augmented Generation):** LangChain (`MemoryVectorStore`), HuggingFace Inference API (Embeddings).

### Схема взаимодействия

```mermaid
graph TD
    subgraph "1. Outlook Client"
        User((Пользователь)) -->|Открывает письмо| Outlook[Microsoft Outlook]
        Outlook -->|Клик "Process Email"| Addin[Next.js + Fluent UI]
        Addin -->|Извлекает данные| OfficeJS(Office.js)
        OfficeJS -->|Отправляет в чат| ChatUI[Vercel AI SDK]
    end

    subgraph "2. Next.js API (Оркестратор)"
        ChatUI <-->|POST /api/chat| APIRoute[API Route]
        APIRoute <-->|Общение & Tool Calling| Model[OpenRouter LLM]
    end

    subgraph "3. MCP Server (Интеграция 1С)"
        APIRoute <-->|MCP Protocol| MCPServer[Node.js MCP Server]
        MCPServer -->|CRUD операции| OneC[(Mock 1C CRM)]
        MCPServer -->|Семантический поиск| VectorDB[(LangChain Vector Store)]
        VectorDB <-->|Векторизация| HFEmbeddings[HuggingFace API]
    end
```

## 📁 Структура проекта

```text
CRMAgentForOutlook/
├── mcp-1c-server/               # Микросервис интеграции с 1С (MCP Server)
│   ├── src/
│   │   └── index.ts             # Логика MCP, RAG (LangChain) и мок-база 1С
│   ├── package.json
│   └── tsconfig.json
├── public/
│   └── manifest.xml             # Манифест надстройки для Outlook (Sideloading)
├── src/
│   └── app/
│       ├── api/chat/route.ts    # Next.js API Route: логика ИИ, системный промпт, инструменты
│       ├── globals.css          # Глобальные стили
│       ├── layout.tsx           # Корневой layout, инъекция Office.js и фикс window.history
│       ├── page.tsx             # Главный UI чата (Fluent UI, рендеринг кнопок и карточек)
│       └── providers.tsx        # Fluent UI Theme Provider
├── .env.local                   # Ключи API (OPENROUTER_API_KEY, HF_TOKEN)
├── next.config.js
├── package.json
└── tailwind.config.ts
```

## 🚀 Запуск для разработки

Для работы системы требуются два параллельно запущенных процесса и загрузка манифеста в Outlook.

### 1. Настройка окружения
Создайте файл `.env.local` в корне проекта и добавьте ключи:
```env
OPENROUTER_API_KEY=ваш_ключ_openrouter
HF_TOKEN=ваш_ключ_huggingface
```

### 2. Запуск MCP-сервера 1С
В первом терминале:
```bash
cd mcp-1c-server
npm install
npm run build && node dist/index.js
```

### 3. Запуск Next.js (Фронтенд)
Во втором терминале (в корне проекта):
```bash
npm install
npm run dev
```
*Сервер запустится на `https://localhost:3000`.*

### 4. Установка в Outlook (Sideloading)
1. Скопируйте файл `public/manifest.xml` в расшаренную сетевую папку (например, `\\localhost\OutlookManifests`).
2. В десктопном Outlook 2021 перейдите в **Файл -> Управление надстройками** (или Центр управления безопасностью -> Надстройки -> Доверенные каталоги).
3. Добавьте сетевую папку как доверенный каталог.
4. Перезапустите Outlook, перейдите в "Получить надстройки" -> "Общая папка" и добавьте `AI CRM Agent`.
5. Откройте любое письмо — на верхней панели появится кнопка **Process Email**.

## 🔮 Пути дальнейшего развития (Production)

1. **Реальная интеграция с 1С:** Замена массива `mockDeals` в `mcp-1c-server` на реальные HTTP/OData запросы к REST API 1С Предприятие.
2. **Персистентная векторная БД:** Замена `MemoryVectorStore` на полноценную базу данных (PgVector, Pinecone, Qdrant) для хранения десятков тысяч сделок.
3. **Автоматическая очистка контекста:** Доработка слушателя `ItemChanged` в `page.tsx` для автоматического сброса истории чата при переключении пользователем на другое письмо.
4. **Аутентификация:** Добавление SSO (Single Sign-On) через Microsoft Entra ID (Azure AD), чтобы ИИ знал, какой именно менеджер сейчас работает со сделкой.