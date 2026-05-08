import { useState, useCallback, useRef } from "react";
import JSZip from "jszip";
import { diff_match_patch } from "diff-match-patch";
import { saveAs } from "file-saver";
import { generateAndDownload } from "./reportGenerator";
import * as pdfjsLib from "pdfjs-dist";
import PdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
pdfjsLib.GlobalWorkerOptions.workerSrc = PdfWorker;

// ── 파일 파서 ─────────────────────────────────────────────────
async function readZip(file) {
  const buf = await file.arrayBuffer();
  return await JSZip.loadAsync(buf);
}

async function extractChunks(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".hwp"))  return await parseHwp(file);
  if (name.endsWith(".hwpx")) return await parseHwpx(file);
  if (name.endsWith(".docx")) return await parseDocx(file);
  if (name.endsWith(".pptx")) return await parsePptx(file);
  if (name.endsWith(".xlsx")) return await parseXlsx(file);
  if (name.endsWith(".pdf"))  return await parsePdf(file);
  throw new Error("지원하지 않는 파일 형식입니다.");
}

async function parseHwp(file) {
  const { parseHwp: kordocParseHwp } = await import("kordoc");
  const buf = await file.arrayBuffer();
  const result = await kordocParseHwp(buf);
  if (!result.success) throw new Error(`HWP 파싱 실패: ${result.error}`);
  const chunks = [];
  let paraNum = 0;
  for (const block of result.blocks) {
    const text = extractBlockText(block).trim();
    if (text) { paraNum++; chunks.push({ text, location: `${paraNum}번째 단락` }); }
  }
  return chunks;
}

function extractBlockText(block) {
  if (block.type === "paragraph" || block.type === "heading" || block.type === "list") {
    const own = block.text ?? "";
    const child = (block.children ?? []).map(extractBlockText).join(" ");
    return [own, child].filter(Boolean).join(" ");
  }
  if (block.type === "table" && block.table) {
    return block.table.cells.flatMap(row => row.map(cell => cell.text ?? "")).filter(Boolean).join(" ");
  }
  return "";
}

async function parseDocx(file) {
  const zip = await readZip(file);
  if (!zip.files["word/document.xml"]) throw new Error("word/document.xml 없음");
  const xml = await zip.files["word/document.xml"].async("string");
  const paraMatches = [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)];
  const chunks = []; let paraNum = 0;
  for (const m of paraMatches) {
    const ts = [...m[0].matchAll(/<w:t[^>]*>(.*?)<\/w:t>/g)].map((t) => t[1]);
    const text = ts.join("").trim();
    if (text) { paraNum++; chunks.push({ text, location: `${paraNum}번째 단락` }); }
  }
  return chunks;
}

async function parseHwpx(file) {
  const zip = await readZip(file);
  const sections = Object.keys(zip.files)
    .filter((k) => k.startsWith("Contents/section") && k.endsWith(".xml")).sort();
  const chunks = []; let paraNum = 0;
  for (const secFile of sections) {
    const xml = await zip.files[secFile].async("string");
    const paraMatches = [...xml.matchAll(/<hp:p[ >][\s\S]*?<\/hp:p>/g)];
    for (const m of paraMatches) {
      const ts = [...m[0].matchAll(/<hp:t\b[^>]*>([^<]*)<\/hp:t>/g)].map((t) => t[1]);
      const text = ts.join("").trim();
      if (text) { paraNum++; chunks.push({ text, location: `${paraNum}번째 단락` }); }
    }
  }
  return chunks;
}


async function parsePptx(file) {
  const zip = await readZip(file);
  const slides = Object.keys(zip.files)
    .filter((k) => k.startsWith("ppt/slides/slide") && k.endsWith(".xml"))
    .sort((a, b) => parseInt(a.match(/slide(\d+)/)?.[1] || 0) - parseInt(b.match(/slide(\d+)/)?.[1] || 0));
  const chunks = [];
  for (let i = 0; i < slides.length; i++) {
    const xml = await zip.files[slides[i]].async("string");
    const paraMatches = [...xml.matchAll(/<a:p[ >][\s\S]*?<\/a:p>/g)];
    let paraNum = 0;
    for (const m of paraMatches) {
      const ts = [...m[0].matchAll(/<a:t[^>]*>(.*?)<\/a:t>/g)].map((t) => t[1]);
      const text = ts.join("").trim();
      if (text) {
        paraNum++;
        chunks.push({ text, location: `슬라이드 ${i + 1} · ${paraNum}번째 단락` });
      }
    }
  }
  return chunks;
}

async function parseXlsx(file) {
  const zip = await readZip(file);
  const sheetNames = {};
  if (zip.files["xl/workbook.xml"]) {
    const wbXml = await zip.files["xl/workbook.xml"].async("string");
    [...wbXml.matchAll(/sheetId="(\d+)"[^>]*name="([^"]+)"/g)]
      .forEach((m) => { sheetNames[m[1]] = m[2]; });
  }
  let shared = [];
  if (zip.files["xl/sharedStrings.xml"]) {
    const xml = await zip.files["xl/sharedStrings.xml"].async("string");
    shared = [...xml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((m) => m[1]);
  }
  const sheets = Object.keys(zip.files)
    .filter((k) => k.startsWith("xl/worksheets/sheet") && k.endsWith(".xml")).sort();
  const chunks = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheetNum = sheets[i].match(/sheet(\d+)/)?.[1] || String(i + 1);
    const sheetName = sheetNames[sheetNum] || `Sheet${sheetNum}`;
    const xml = await zip.files[sheets[i]].async("string");
    const cellTexts = [];
    [...xml.matchAll(/<c[^>]*t="s"[^>]*><v>(\d+)<\/v>/g)]
      .forEach((m) => { const idx = parseInt(m[1]); if (shared[idx]) cellTexts.push(shared[idx]); });
    [...xml.matchAll(/<t[^>]*>([^<]+)<\/t>/g)]
      .forEach((m) => { if (m[1].trim()) cellTexts.push(m[1]); });
    const text = cellTexts.join(" ").trim();
    if (text) chunks.push({ text, location: sheetName });
  }
  return chunks;
}

async function parsePdf(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const chunks = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    // 같은 y좌표 기준으로 줄 묶기
    const lines = {};
    for (const item of content.items) {
      if (!item.str.trim()) continue;
      const y = Math.round(item.transform[5]);
      if (!lines[y]) lines[y] = [];
      lines[y].push(item.str);
    }
    // y좌표 내림차순 정렬 (위→아래)
    const sorted = Object.keys(lines)
      .sort((a, b) => b - a)
      .map((y) => lines[y].join("").trim())
      .filter((t) => t);
    sorted.forEach((text, idx) => {
      chunks.push({ text, location: `${i}페이지 · ${idx + 1}번째 줄` });
    });
  }
  return chunks;
}


// ── 텍스트 정규화 ─────────────────────────────────────────────
function normalizeText(text) {
  return text
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2015]/g, "-")
    .replace(/[\u00A0\u2003\u2002\u2001]/g, " ")
    .replace(/\r\n/g, "\n")
    .replace(/\s*([-·/])\s*/g, "$1")
    .replace(/\s+([.,!?])/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

// ── 유사도 계산 ───────────────────────────────────────────────
function similarity(a, b) {
  if (!a || !b) return 0;
  const dmp = new diff_match_patch();
  const diffs = dmp.diff_main(normalizeText(a), normalizeText(b));
  dmp.diff_cleanupSemantic(diffs);
  let common = 0, total = 0;
  for (const [op, text] of diffs) {
    total += text.length;
    if (op === 0) common += text.length;
  }
  return total === 0 ? 1 : common / total;
}

// ── 단락 매칭 ─────────────────────────────────────────────────
function matchChunks(origChunks, modChunks, threshold = 0.5) {
  const n = origChunks.length;
  const m = modChunks.length;

  // 1. 유사도 행렬 계산 (O(n×m))
  const sim = Array.from({ length: n }, (_, i) =>
    Array.from({ length: m }, (_, j) =>
      similarity(origChunks[i].text, modChunks[j].text)
    )
  );

  // 2. DP 테이블: 순서를 보존하면서 유사도 합을 최대화
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      if (sim[i - 1][j - 1] >= threshold) {
        dp[i][j] = Math.max(dp[i][j], dp[i - 1][j - 1] + sim[i - 1][j - 1]);
      }
    }
  }

  // 3. 역추적으로 매칭 쌍 복원
  const matchedPairs = new Map();
  let i = n, j = m;
  while (i > 0 && j > 0) {
    const canMatch = sim[i - 1][j - 1] >= threshold;
    if (canMatch && dp[i][j] === dp[i - 1][j - 1] + sim[i - 1][j - 1]) {
      matchedPairs.set(i - 1, j - 1);
      i--; j--;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      i--;
    } else {
      j--;
    }
  }

  // 4. 결과 조합
  const usedMod = new Set(matchedPairs.values());
  const result = [];

  for (let oi = 0; oi < n; oi++) {
    if (matchedPairs.has(oi)) {
      const mi = matchedPairs.get(oi);
      result.push({ orig: origChunks[oi], mod: modChunks[mi], score: sim[oi][mi] });
    } else {
      result.push({ orig: origChunks[oi], mod: null, score: 0 });
    }
  }

  for (let mi = 0; mi < m; mi++) {
    if (!usedMod.has(mi)) {
      result.push({ orig: null, mod: modChunks[mi], score: 0 });
    }
  }

  return result;
}

// ── Inline diff ───────────────────────────────────────────────
function buildInlineDiff(origText, modText) {
  const dmp = new diff_match_patch();
  const diffs = dmp.diff_main(normalizeText(origText), normalizeText(modText));
  dmp.diff_cleanupSemantic(diffs);
  return diffs;
}

function classifyPair(pair) {
  if (!pair.orig) return "added";
  if (!pair.mod)  return "deleted";
  if (pair.score === 1) return "same";
  return "changed";
}

async function compare(origFile, modFile) {
  const [origChunks, modChunks] = await Promise.all([
    extractChunks(origFile), extractChunks(modFile),
  ]);
  const pairs = matchChunks(origChunks, modChunks);
  return pairs.map((pair) => ({
    ...pair,
    type: classifyPair(pair),
    diffs: pair.orig && pair.mod ? buildInlineDiff(pair.orig.text, pair.mod.text) : null,
  }));
}


// ── InlineDiff 렌더 ───────────────────────────────────────────
function InlineDiff({ diffs }) {
  return (
    <span>
      {diffs.map(([op, text], i) => {
        if (op === 0)  return <span key={i}>{text}</span>;
        if (op === -1) return <mark key={i} style={{ background: "#FEE2E2", color: "#B91C1C", borderRadius: 3, padding: "0 2px", textDecoration: "line-through" }}>{text}</mark>;
        if (op === 1)  return <mark key={i} style={{ background: "#DCFCE7", color: "#166534", borderRadius: 3, padding: "0 2px" }}>{text}</mark>;
        return null;
      })}
    </span>
  );
}

// ── 파일 업로드 박스 ──────────────────────────────────────────
function UploadBox({ label, file, onFile, color }) {
  const ref = useRef();
  const onDrop = (e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) onFile(f); };
  return (
    <div
      onDrop={onDrop} onDragOver={(e) => e.preventDefault()}
      onClick={() => ref.current.click()}
      style={{ border: `2px dashed ${file ? color : "#CBD5E1"}`, borderRadius: 14, padding: "28px 20px", textAlign: "center", cursor: "pointer", background: file ? color + "10" : "#fff", transition: "all 0.2s", flex: 1 }}
      onMouseEnter={(e) => { e.currentTarget.style.borderColor = color; }}
      onMouseLeave={(e) => { e.currentTarget.style.borderColor = file ? color : "#CBD5E1"; }}
    >
      <input ref={ref} type="file" accept=".hwpx,.docx,.pptx,.xlsx,.pdf" style={{ display: "none" }} onChange={(e) => { if (e.target.files[0]) onFile(e.target.files[0]); }} />
      <div style={{ fontSize: 28, marginBottom: 8 }}>{file ? "📄" : "⬆️"}</div>
      <div style={{ fontWeight: 700, fontSize: 13, color: file ? color : "#64748B", marginBottom: 4 }}>{label}</div>
      {file
        ? <div style={{ fontSize: 12, color: "#64748B" }}>{file.name}<br />{(file.size / 1024).toFixed(1)} KB</div>
        : <div style={{ fontSize: 12, color: "#94A3B8" }}>HWPX · DOCX · PPTX · XLSX · PDF</div>}
    </div>
  );
}

// ── 메인 컴포넌트 ─────────────────────────────────────────────
export default function Comparator() {
  const [origFile,   setOrigFile]   = useState(null);
  const [modFile,    setModFile]    = useState(null);
  const [loading,    setLoading]    = useState(false);
  const [generating, setGenerating] = useState(false);
  const [results,    setResults]    = useState(null);
  const [error,      setError]      = useState(null);
  const [filter,     setFilter]     = useState("changed");

  const handleCompare = useCallback(async () => {
    if (!origFile || !modFile) return;
    setLoading(true); setError(null); setResults(null);
    try {
      const pairs = await compare(origFile, modFile);
      setResults(pairs);
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, [origFile, modFile]);

const handleGenerate = useCallback(async () => {
  if (!results) return;
  setGenerating(true);
  try {
    const changedPairs = results
      .filter((r) => r.type === "changed" || r.type === "deleted")
      .map((r) => ({
        location: r.orig?.location || r.mod?.location || "",
        origText: r.orig?.text || "",
        modText:  r.mod?.text  || "[삭제]",
      }));
    await generateAndDownload(origFile, changedPairs);
    alert("조치내역서가 다운로드되었습니다.\n다운로드 폴더를 확인해 주세요.");
  } catch (e) { setError(e.message); }
  finally { setGenerating(false); }
}, [results, origFile]);

  const counts = results ? {
    same:    results.filter((r) => r.type === "same").length,
    changed: results.filter((r) => r.type === "changed").length,
    added:   results.filter((r) => r.type === "added").length,
    deleted: results.filter((r) => r.type === "deleted").length,
  } : null;

  const filtered = results?.filter((r) => filter === "all" ? true : r.type === filter);

  const TYPE_LABEL = { same: "변경 없음", changed: "변경됨", added: "추가됨", deleted: "삭제됨" };
  const TYPE_COLOR = {
    same:    { bg: "#F0FDF4", border: "#86EFAC", text: "#166534", badge: "#16A34A" },
    changed: { bg: "#FEF3C7", border: "#FCD34D", text: "#92400E", badge: "#D97706" },
    added:   { bg: "#EFF6FF", border: "#93C5FD", text: "#1E40AF", badge: "#2563EB" },
    deleted: { bg: "#FEF2F2", border: "#FCA5A5", text: "#991B1B", badge: "#DC2626" },
  };

  return (
    <div style={{ fontFamily: "'Pretendard', 'Apple SD Gothic Neo', sans-serif", minHeight: "100vh", background: "#F8FAFC", color: "#0F172A" }}>
      <div style={{ background: "#1B2E5E", color: "#fff", padding: "20px 28px", display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ width: 36, height: 36, background: "#7C3AED", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>⚖️</div>
        <div>
          <div style={{ fontWeight: 700, fontSize: 16 }}>비식별화 비교기 + 조치내역서 생성</div>
          <div style={{ fontSize: 12, color: "#94A3B8", marginTop: 2 }}>원본과 수정본을 비교하고 조치내역서를 자동 생성합니다</div>
        </div>
      </div>

      <div style={{ maxWidth: 960, margin: "0 auto", padding: "28px 20px" }}>
        <div style={{ display: "flex", gap: 16, marginBottom: 20 }}>
          <UploadBox label="원본 파일" file={origFile} onFile={setOrigFile} color="#DC2626" />
          <div style={{ display: "flex", alignItems: "center", fontSize: 24, color: "#CBD5E1", flexShrink: 0 }}>→</div>
          <UploadBox label="수정본 파일" file={modFile} onFile={setModFile} color="#16A34A" />
        </div>

        <button onClick={handleCompare} disabled={!origFile || !modFile || loading}
          style={{ width: "100%", padding: "14px", borderRadius: 12, border: "none", background: origFile && modFile && !loading ? "#1B2E5E" : "#CBD5E1", color: "#fff", fontWeight: 700, fontSize: 15, cursor: origFile && modFile && !loading ? "pointer" : "not-allowed", marginBottom: 24 }}>
          {loading ? "⏳ 비교 중..." : "비교 시작"}
        </button>

        {error && (
          <div style={{ background: "#FEE2E2", border: "1px solid #EF4444", borderRadius: 10, padding: "14px 18px", color: "#B91C1C", marginBottom: 16 }}>⚠️ {error}</div>
        )}

        {results && !loading && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 20 }}>
              {[
                { key: "changed", label: "변경됨",    icon: "✏️" },
                { key: "added",   label: "추가됨",    icon: "➕" },
                { key: "deleted", label: "삭제됨",    icon: "🗑️" },
                { key: "same",    label: "변경 없음", icon: "✅" },
              ].map(({ key, label, icon }) => {
                const c = TYPE_COLOR[key];
                return (
                  <div key={key} style={{ background: c.bg, border: `1px solid ${c.border}`, borderRadius: 12, padding: "16px 18px" }}>
                    <div style={{ fontSize: 13, color: c.text, fontWeight: 600 }}>{icon} {label}</div>
                    <div style={{ fontSize: 30, fontWeight: 800, color: c.badge, marginTop: 4 }}>{counts[key]}</div>
                    <div style={{ fontSize: 12, color: c.text, opacity: 0.7 }}>단락</div>
                  </div>
                );
              })}
            </div>

            <button onClick={handleGenerate} disabled={generating}
              style={{ width: "100%", padding: "14px", borderRadius: 12, border: "none", background: generating ? "#CBD5E1" : "#16A34A", color: "#fff", fontWeight: 700, fontSize: 15, cursor: generating ? "not-allowed" : "pointer", marginBottom: 20 }}>
              {generating ? "⏳ 생성 중..." : "📄 조치내역서 DOCX 다운로드"}
            </button>

            <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>
              {[
                { key: "changed", label: `변경됨 (${counts.changed})` },
                { key: "added",   label: `추가됨 (${counts.added})` },
                { key: "deleted", label: `삭제됨 (${counts.deleted})` },
                { key: "all",     label: `전체 (${results.length})` },
              ].map(({ key, label }) => {
                const active = filter === key;
                const c = key === "all" ? { badge: "#1B2E5E" } : TYPE_COLOR[key];
                return (
                  <button key={key} onClick={() => setFilter(key)}
                    style={{ padding: "6px 16px", borderRadius: 20, border: "none", cursor: "pointer", background: active ? c.badge : "#E2E8F0", color: active ? "#fff" : "#475569", fontWeight: active ? 700 : 400, fontSize: 13 }}>
                    {label}
                  </button>
                );
              })}
            </div>

            {filtered.length === 0
              ? <div style={{ textAlign: "center", padding: "40px", color: "#94A3B8" }}>해당 항목이 없습니다.</div>
              : filtered.map((pair, i) => {
                  const c = TYPE_COLOR[pair.type];
                  const location = pair.orig?.location || pair.mod?.location || "";
                  return (
                    <div key={i} style={{ background: "#fff", borderRadius: 12, border: `1px solid ${c.border}`, marginBottom: 10, overflow: "hidden" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 16px", background: c.bg, borderBottom: `1px solid ${c.border}` }}>
                        <span style={{ background: c.badge, color: "#fff", borderRadius: 6, padding: "2px 8px", fontSize: 12, fontWeight: 700 }}>{TYPE_LABEL[pair.type]}</span>
                        <span style={{ background: "#EEF2FF", color: "#3730A3", borderRadius: 5, padding: "2px 8px", fontSize: 12, fontWeight: 600 }}>{location}</span>
                        {pair.type === "changed" && (
                          <span style={{ fontSize: 12, color: "#94A3B8", marginLeft: "auto" }}>유사도 {Math.round(pair.score * 100)}%</span>
                        )}
                      </div>
                      <div style={{ padding: "12px 16px", fontSize: 14, lineHeight: 1.8 }}>
                        {pair.type === "changed" && pair.diffs && <InlineDiff diffs={pair.diffs} />}
                        {pair.type === "same"    && <span style={{ color: "#64748B" }}>{pair.orig.text}</span>}
                        {pair.type === "deleted" && <span style={{ color: "#B91C1C", textDecoration: "line-through" }}>{pair.orig.text}</span>}
                        {pair.type === "added"   && <span style={{ color: "#166534" }}>{pair.mod.text}</span>}
                      </div>
                    </div>
                  );
                })
            }

            <div style={{ marginTop: 20, background: "#F8FAFC", border: "1px solid #CBD5E1", borderRadius: 12, padding: "14px 18px", fontSize: 13, color: "#64748B", lineHeight: 1.7 }}>
              <strong style={{ color: "#334155" }}>ℹ️ 안내</strong><br />
              · <mark style={{ background: "#FEE2E2", color: "#B91C1C", borderRadius: 3, padding: "0 2px" }}>빨간 텍스트</mark>는 원본에서 삭제된 내용,{" "}
              <mark style={{ background: "#DCFCE7", color: "#166534", borderRadius: 3, padding: "0 2px" }}>초록 텍스트</mark>는 수정본에서 추가된 내용입니다.<br />
              · 조치내역서 다운로드 후 문서번호·부서·담당자·비식별화 방법 등을 직접 입력해 주세요.<br />
              · 단락이 추가·삭제된 경우 유사도 기반으로 매칭하므로 일부 오차가 있을 수 있습니다.
            </div>
          </>
        )}
      </div>
    </div>
  );
}
