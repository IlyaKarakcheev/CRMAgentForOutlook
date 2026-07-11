"use client";

import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import React from "react";

export function Providers({ children }: { children: React.ReactNode }) {
  return (
    <FluentProvider theme={webLightTheme} style={{ height: '100%' }}>
      {children}
    </FluentProvider>
  );
}