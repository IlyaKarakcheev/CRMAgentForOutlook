import type { Metadata, Viewport } from "next";
import { Providers } from "./providers";
import "./globals.css";

export const metadata: Metadata = {
  title: "Outlook CRM AI Agent",
  description: "AI Agent for 1C CRM directly inside Outlook",
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <head>
        {/* Сохраняем оригинальные методы history, так как Office.js может их затирать вне Outlook */}
        <script
          dangerouslySetInnerHTML={{
            __html: `
              window._originalHistoryPushState = window.history.pushState;
              window._originalHistoryReplaceState = window.history.replaceState;
            `,
          }}
        />
        <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>
        <script
          dangerouslySetInnerHTML={{
            __html: `
              if (!window.history.replaceState) {
                window.history.replaceState = window._originalHistoryReplaceState;
              }
              if (!window.history.pushState) {
                window.history.pushState = window._originalHistoryPushState;
              }
            `,
          }}
        />
      </head>
      <body>
        <Providers>{children}</Providers>
      </body>
    </html>
  );
}