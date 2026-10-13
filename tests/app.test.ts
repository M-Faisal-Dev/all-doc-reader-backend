import { afterEach, describe, expect, test } from "bun:test";
import type { Server } from "node:http";
import { createApp } from "../src/app";
const servers: Server[] = [];
afterEach(() => { for (const server of servers.splice(0)) server.close(); });
function setup(fetcher: typeof fetch) {
  const server = createApp({ gotenbergUrl: "http://localhost:3005", origins: ["http://localhost:8081"], fetcher }).listen(0);
  servers.push(server);
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("Missing address");
  return `http://localhost:${address.port}`;
}
const fake = (handler: (request: string | URL | Request, init?: RequestInit) => Promise<Response>) => handler as typeof fetch;
function form(name = "sample.docx", data = new Uint8Array([0x50, 0x4b, 3, 4, 1, 2, 3, 4])) { const body = new FormData(); body.append("files", new Blob([data]), name); return body; }
describe("conversion API", () => {
  test("health reports upstream availability", async () => {
    const url = setup(fake(async () => new Response("{}")));
    expect((await (await fetch(`${url}/health`)).json()).gotenberg).toBe("up");
  });
  test("converts a Word upload and returns PDF with local CORS", async () => {
    const url = setup(fake(async (request, init) => {
      expect(String(request)).toEndWith("/forms/libreoffice/convert");
      expect((init?.body as FormData).get("files")).toBeInstanceOf(File);
      return new Response("%PDF-1.7\nfixture", { headers: { "Content-Type": "application/pdf" } });
    }));
    const result = await fetch(`${url}/api/v1/convert/word`, { method: "POST", body: form(), headers: { Origin: "http://localhost:8081" } });
    expect(result.status).toBe(200);
    expect(result.headers.get("access-control-allow-origin")).toBe("http://localhost:8081");
    expect(await result.text()).toStartWith("%PDF-");
  });
  test("rejects unsupported files without contacting Gotenberg", async () => {
    let calls = 0;
    const url = setup(fake(async () => { calls++; return new Response(); }));
    const result = await fetch(`${url}/api/v1/convert/word`, { method: "POST", body: form("payload.exe") });
    expect(result.status).toBe(415); expect(calls).toBe(0);
  });
  test("rejects missing uploads", async () => {
    const url = setup(fake(async () => new Response()));
    expect((await fetch(`${url}/api/v1/convert/word`, { method: "POST" })).status).toBe(400);
  });
  test("reports upstream errors and rejects non-PDF responses", async () => {
    const badFile = setup(fake(async () => new Response("bad", { status: 400 })));
    expect((await fetch(`${badFile}/api/v1/convert/word`, { method: "POST", body: form() })).status).toBe(422);
    const badResponse = setup(fake(async () => new Response("not a pdf")));
    expect((await fetch(`${badResponse}/api/v1/convert/word`, { method: "POST", body: form() })).status).toBe(502);
  });
});

describe("Excel conversion API", () => {
  for (const extension of ["xls", "xlsx"]) {
    test(`converts ${extension} without forcing sheets onto one page`, async () => {
      const url = setup(fake(async (_request, init) => {
        const body = init?.body as FormData;
        const file = body.get("files") as File;
        expect(file.name).toBe(`document.${extension}`);
        expect(file.type).toBe(extension === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel");
        expect(body.has("singlePageSheets")).toBe(false);
        return new Response("%PDF-1.7\nfixture");
      }));
      const data = extension === "xls" ? new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]) : undefined;
      const result = await fetch(`${url}/api/v1/convert/excel`, { method: "POST", body: form(`sample.${extension}`, data) });
      expect(result.status).toBe(200);
      expect(await result.text()).toStartWith("%PDF-");
    });
  }
  test("rejects wrong route types and invalid Excel signatures", async () => {
    let calls = 0;
    const url = setup(fake(async () => { calls++; return new Response(); }));
    for (const [route, body] of [
      ["excel", form("sample.docx")],
      ["word", form("sample.xlsx")],
      ["excel", form("sample.xlsx", new Uint8Array([1,2,3]))],
      ["excel", form("sample.xls")],
    ] as const) {
      expect((await fetch(`${url}/api/v1/convert/${route}`, { method: "POST", body })).status).toBe(415);
    }
    expect(calls).toBe(0);
  });
});

describe("PowerPoint conversion API", () => {
  for (const extension of ["ppt", "pptx"]) {
    test(`converts ${extension} using the presentation MIME type`, async () => {
      const url = setup(fake(async (_request, init) => {
        const file = (init?.body as FormData).get("files") as File;
        expect(file.name).toBe(`document.${extension}`);
        expect(file.type).toBe(extension === "pptx" ? "application/vnd.openxmlformats-officedocument.presentationml.presentation" : "application/vnd.ms-powerpoint");
        return new Response("%PDF-1.7\nfixture");
      }));
      const data = extension === "ppt" ? new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]) : undefined;
      const result = await fetch(`${url}/api/v1/convert/powerpoint`, { method: "POST", body: form(`slides.${extension}`, data) });
      expect(result.status).toBe(200);
      expect(await result.text()).toStartWith("%PDF-");
    });
  }
  test("rejects wrong types and damaged PowerPoint signatures", async () => {
    let calls = 0;
    const url = setup(fake(async () => { calls++; return new Response(); }));
    for (const [route, body] of [
      ["powerpoint", form("sample.docx")],
      ["word", form("slides.pptx")],
      ["excel", form("slides.pptx")],
      ["powerpoint", form("slides.pptx", new Uint8Array([1,2,3]))],
      ["powerpoint", form("slides.ppt")],
    ] as const) {
      expect((await fetch(`${url}/api/v1/convert/${route}`, { method: "POST", body })).status).toBe(415);
    }
    expect(calls).toBe(0);
  });
});
