import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Compara — Ventas, sin vueltas",
  description: "Explorá tus ventas desde cualquier Excel, de forma clara y segura.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
