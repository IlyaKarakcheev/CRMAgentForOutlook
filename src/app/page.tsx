"use client";

import React, { useEffect, useState, useRef } from "react";
import { useChat } from "ai/react";
import {
  Button,
  Input,
  Text,
  Spinner,
  Card,
  CardHeader,
  CardFooter,
  Badge,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import { SendRegular, MailRegular, BotRegular, PersonRegular, SearchRegular, DocumentAddRegular, DeleteRegular, EditRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  header: {
    padding: "12px 16px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    boxShadow: tokens.shadow4,
    zIndex: 10,
  },
  chatArea: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  messageUser: {
    alignSelf: "flex-end",
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke2}`,
    borderRadius: "12px 12px 0 12px",
    padding: "12px",
    maxWidth: "85%",
  },
  messageBot: {
    alignSelf: "flex-start",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "12px 12px 12px 0",
    padding: "12px",
    maxWidth: "85%",
    boxShadow: tokens.shadow2,
  },
  toolCard: {
    backgroundColor: tokens.colorNeutralBackground1Hover,
    border: `1px dashed ${tokens.colorNeutralStroke2}`,
  },
  inputArea: {
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    display: "flex",
    gap: "8px",
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: "16px",
    color: tokens.colorNeutralForeground2,
  },
  btnSelected: {
    backgroundColor: tokens.colorPaletteGreenBackground3,
    color: tokens.colorNeutralForegroundInverted,
    borderColor: "transparent",
    cursor: "default",
    ":hover": {
      backgroundColor: tokens.colorPaletteGreenBackground3,
      color: tokens.colorNeutralForegroundInverted,
    },
    ":active": {
      backgroundColor: tokens.colorPaletteGreenBackground3,
      color: tokens.colorNeutralForegroundInverted,
    }
  }
});

export default function Home() {
  const styles = useStyles();
  const [isOfficeInitialized, setIsOfficeInitialized] = useState(false);
  const [emailData, setEmailData] = useState<{ subject: string; from: string; body: string } | null>(null);

  const { messages, input, handleInputChange, handleSubmit, append, isLoading } = useChat({
    api: '/api/chat'
  });

  useEffect(() => {
    if (typeof window !== "undefined" && (window as any).Office) {
      (window as any).Office.onReady((info: any) => {
        if (info.host === (window as any).Office.HostType.Outlook) {
          setIsOfficeInitialized(true);
          
          // Безопасная регистрация обработчика ItemChanged с блоком try-catch
          try {
            const office = (window as any).Office;
            if (office && office.context && office.context.mailbox && office.context.mailbox.addHandlerAsync) {
              office.context.mailbox.addHandlerAsync(
                office.EventType.ItemChanged,
                function (eventArgs: any) {
                  // Вызывается при переключении писем, когда панель закреплена (Pinned)
                  console.log("Email item changed.");
                }
              );
            }
          } catch (e) {
            console.error("Failed to register ItemChanged handler", e);
          }
        }
      });
    }
  }, []);

  const messagesEndRef = useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, isLoading]);

  const processEmail = () => {
    const office = (window as any).Office;
    if (!office || !office.context || !office.context.mailbox || !office.context.mailbox.item) {
      alert("Office.js is not fully loaded or no email item is open. To test in browser, just type a message below!");
      return;
    }

    const item = office.context.mailbox.item;
    const subject = item.subject || "No Subject";
    const from = item.from?.emailAddress || "Unknown Sender";

    item.body.getAsync("text", (result: any) => {
      if (result.status === office.AsyncResultStatus.Succeeded) {
        const body = result.value;
        setEmailData({ subject, from, body });
        append({
          role: "user",
          content: `Обработать письмо.\n\nОт: ${from}\nТема: ${subject}\n\nТекст письма:\n${body}`
        });
      } else {
        alert("Ошибка при чтении тела письма.");
      }
    });
  };

  const renderToolInvocation = (toolInvocation: any) => {
    const { toolCallId, toolName, state, result } = toolInvocation;
    
    let icon = <BotRegular />;
    let title = "Использование инструмента";
    
    if (toolName === 'search_deal_exact') {
      icon = <SearchRegular />;
      title = "Точный поиск сделки в 1С";
    } else if (toolName === 'search_similar_deals_rag') {
      icon = <SearchRegular />;
      title = "Поиск похожих сделок в 1С";
    } else if (toolName === 'create_deal') {
      icon = <DocumentAddRegular />;
      title = "Создание сделки в 1С";
    } else if (toolName === 'link_deal') {
      icon = <BotRegular />;
      title = "Привязка письма к сделке";
    } else if (toolName === 'delete_deal') {
      icon = <DeleteRegular />;
      title = "Удаление сделки из 1С";
    } else if (toolName === 'update_deal_amount') {
      icon = <EditRegular />;
      title = "Обновление суммы сделки";
    }

    return (
      <Card key={toolCallId} className={styles.toolCard} size="small">
        <CardHeader
          image={icon}
          header={<Text weight="semibold" size={200}>{title}</Text>}
          description={<Text size={100} color="neutral-secondary">{state === 'result' ? "Завершено" : "Выполняется..."}</Text>}
        />
        {state === 'result' && result && (
          <CardFooter>
            {result.success ? (
              <MessageBar intent="success" layout="multiline" shape="square">
                <MessageBarBody>
                  <MessageBarTitle>{result.message || "Успешно"}</MessageBarTitle>
                  {result.deal && <Text block size={200}>Сделка: {result.deal.client} ({result.deal.amount} руб.)</Text>}
                  {result.deals && result.deals.map((d: any) => <Text key={d.id} block size={200}>• {d.client}: {d.description}</Text>)}
                </MessageBarBody>
              </MessageBar>
            ) : (
              <MessageBar intent="warning" shape="square">
                <MessageBarBody>{result.message}</MessageBarBody>
              </MessageBar>
            )}
          </CardFooter>
        )}
      </Card>
    );
  };

  const renderQuickActions = (message: any, index: number, allMessages: any[]) => {
    if (message.role !== 'assistant') return null;

    const ragInvocation = message.toolInvocations?.find((t: any) => t.toolName === 'search_similar_deals_rag');
    if (!ragInvocation || !ragInvocation.result || !ragInvocation.result.success) {
      return null;
    }

    const deals = ragInvocation.result.deals || [];
    let selectedAction: string | null = null;

    // Ищем выбор пользователя после этого сообщения
    for (let i = index + 1; i < allMessages.length; i++) {
      const m = allMessages[i];
      if (m.role === 'user') {
        const text = m.content.toLowerCase();
        if (text.includes("создадим новую") || text.includes("создать новую")) {
          selectedAction = 'create_new';
          break;
        } else if (text.includes("да, привяжи")) {
          const match = text.match(/id:\s*(\d+)/i);
          if (match) {
            selectedAction = `link_${match[1]}`;
          } else {
            selectedAction = 'link_generic';
          }
          break;
        }
      } else if (m.role === 'assistant' && m.toolInvocations) {
        // Также учитываем, если ИИ сам вызвал инструмент после (если пользователь ответил по-другому)
        if (m.toolInvocations.some((t: any) => t.toolName === 'create_deal')) {
          selectedAction = 'create_new';
          break;
        }
        if (m.toolInvocations.some((t: any) => t.toolName === 'link_deal')) {
          selectedAction = 'link_generic';
          break;
        }
      }
    }

    const isResolved = selectedAction !== null;

    const getButtonProps = (actionId: string, defaultAppearance: any) => {
      if (!isResolved) {
        return { appearance: defaultAppearance, disabled: false };
      }
      if (selectedAction === actionId || (actionId.startsWith('link_') && selectedAction === 'link_generic')) {
        return { 
          appearance: "primary" as const, 
          className: styles.btnSelected,
          disabledFocusable: true, // делает кнопку некликабельной, но не применяет нативные стили disabled
          onClick: undefined
        };
      }
      return { appearance: "outline" as const, disabled: true };
    };

    return (
      <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', marginTop: '4px' }}>
        {deals.map((deal: any) => {
          const actionId = `link_${deal.id}`;
          const props = getButtonProps(actionId, "secondary");
          return (
            <Button 
              key={deal.id}
              {...props}
              size="small"
              onClick={!isResolved ? () => append({ role: 'user', content: `Да, привяжи к найденной сделке ID: ${deal.id}` }) : undefined}
            >
              Привязать к {deal.client} (ID: {deal.id})
            </Button>
          );
        })}
        {deals.length === 0 && (
          <Button 
            {...getButtonProps('link_generic', "secondary")}
            size="small"
            onClick={!isResolved ? () => append({ role: 'user', content: `Да, привяжи к найденной сделке` }) : undefined}
          >
            Привязать к найденной
          </Button>
        )}
        <Button 
          {...getButtonProps('create_new', "primary")}
          size="small"
          onClick={!isResolved ? () => append({ role: 'user', content: `Нет, давай создадим новую сделку` }) : undefined}
        >
          Создать новую
        </Button>
      </div>
    );
  };

  return (
    <div className={styles.container}>
      {/* Header */}
      <div className={styles.header}>
        <Text weight="bold" size={500}>CRM Agent</Text>
        {!isOfficeInitialized ? (
          <Badge color="danger" shape="rounded">Browser Mode</Badge>
        ) : (
          <Badge color="success" shape="rounded">Outlook Mode</Badge>
        )}
      </div>

      {/* Messages */}
      <div className={styles.chatArea}>
        {messages.length === 0 && (
          <div className={styles.emptyState}>
            <MailRegular fontSize={48} />
            <Text align="center">Откройте письмо и нажмите кнопку ниже, чтобы ИИ проанализировал его.</Text>
            <Button appearance="primary" icon={<MailRegular />} onClick={processEmail} size="large">
              Обработать письмо
            </Button>
          </div>
        )}

        {(() => {
          const groupedMessages: { id: string, role: string, messages: any[] }[] = [];
          messages.forEach(m => {
            if (groupedMessages.length > 0 && groupedMessages[groupedMessages.length - 1].role === m.role) {
              groupedMessages[groupedMessages.length - 1].messages.push(m);
            } else {
              groupedMessages.push({ id: m.id, role: m.role, messages: [m] });
            }
          });

          return groupedMessages.map((group) => (
            <div key={group.id} className={group.role === 'user' ? styles.messageUser : styles.messageBot} style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {group.role === 'user' ? <PersonRegular fontSize={16} /> : <BotRegular fontSize={16} />}
                <Text weight="semibold" size={200}>{group.role === 'user' ? 'Вы' : 'CRM Ассистент'}</Text>
              </div>
              
              {group.messages.map((m) => (
                <React.Fragment key={m.id}>
                  {m.content && <Text block style={{ whiteSpace: "pre-wrap" }}>{m.content}</Text>}
                  {m.toolInvocations?.map(renderToolInvocation)}
                </React.Fragment>
              ))}

              {group.messages.map((m) => {
                const globalIndex = messages.findIndex(msg => msg.id === m.id);
                return <React.Fragment key={`qa-${m.id}`}>{renderQuickActions(m, globalIndex, messages)}</React.Fragment>;
              })}
            </div>
          ));
        })()}
        
        {isLoading && (
          <div className={styles.messageBot}>
             <Spinner size="tiny" label="ИИ думает..." />
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input */}
      <div className={styles.inputArea}>
        <form onSubmit={handleSubmit} style={{ display: 'flex', gap: '8px', width: '100%' }}>
          <Input
            value={input}
            onChange={(e, data) => handleInputChange({ target: { value: data.value } } as any)}
            placeholder="Напишите сообщение ИИ..."
            aria-label="Поле ввода сообщения"
            title="Поле ввода сообщения"
            style={{ flexGrow: 1 }}
            disabled={isLoading}
          />
          <Button 
            type="submit" 
            appearance="primary" 
            icon={<SendRegular />}
            aria-label="Отправить"
            title="Отправить"
            disabled={isLoading || !input}
          />
        </form>
      </div>
    </div>
  );
}