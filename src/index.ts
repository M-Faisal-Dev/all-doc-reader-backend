import { createApp } from "./app";
const port = Number(process.env.PORT || 5000);
const host = process.env.HOST || "0.0.0.0";
if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error("Invalid PORT");
const app = createApp({
  gotenbergUrl: process.env.GOTENBERG_URL || "http://127.0.0.1:3005",
  origins: (process.env.CORS_ORIGINS || "http://localhost:8081,http://localhost:8082,http://localhost:19006").split(",").map((origin) => origin.trim()),
  timeoutMs: Number(process.env.CONVERSION_TIMEOUT_MS || 90000),
});
const server = app.listen(port, host, () => console.log(`All Doc Reader backend listening on ${host}:${port}`));
for (const signal of ["SIGINT", "SIGTERM"] as const) process.on(signal, () => { server.close(() => process.exit(0)); setTimeout(() => process.exit(1), 10000).unref(); });
