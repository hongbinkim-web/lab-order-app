import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const KEY_DB = "lab_order_db";
const KEY_HISTORY = "lab_order_history";

const load = (key, fallback) => {
  try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; }
};

const SAMPLE_DB = [
  { id: 1, name: "RPMI 1640 배지", cat: "R8758", maker: "Sigma-Aldrich", link: "https://www.sigmaaldrich.com" },
  { id: 2, name: "FBS (Fetal Bovine Serum)", cat: "16000044", maker: "Gibco", link: "" },
  { id: 3, name: "Trypan Blue Solution", cat: "T8154", maker: "Sigma-Aldrich", link: "" },
];

function Notification({ msg }) {
  if (!msg) return null;
  return (
    <div style={{ position:"fixed", bottom:24, right:24, background:"#1a1a1a", color:"#fff", padding:"10px 18px", borderRadius:9, fontSize:13, zIndex:999, boxShadow:"0 4px 16px rgba(0,0,0,0.18)" }}>
      {msg}
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState("order");
  const [db, setDb] = useState(() => load(KEY_DB, SAMPLE_DB));
  const [history, setHistory] = useState(() => load(KEY_HISTORY, []));
  const [selected, setSelected] = useState({});
  const [search, setSearch] = useState("");
  const [notif, setNotif] = useState("");
  const [form, setForm] = useState({ name:"", cat:"", maker:"", link:"" });
  const [gmailOpen, setGmailOpen] = useState(false);
  const [gmailTo, setGmailTo] = useState("");
  const [gmailSubject, setGmailSubject] = useState("");
  const [nextId, setNextId] = useState(() => {
    const saved = load(KEY_DB, SAMPLE_DB);
    return saved.length ? Math.max(...saved.map(x => x.id)) + 1 : 100;
  });
  const [mapModal, setMapModal] = useState(null);

  const notify = useCallback((msg) => {
    setNotif(msg); setTimeout(() => setNotif(""), 2400);
  }, []);

  useEffect(() => { localStorage.setItem(KEY_DB, JSON.stringify(db)); }, [db]);
  useEffect(() => { localStorage.setItem(KEY_HISTORY, JSON.stringify(history)); }, [history]);

  const addItem = () => {
    if (!form.name.trim() || !form.cat.trim()) { notify("품목명과 카탈로그번호는 필수입니다."); return; }
    setDb(prev => [...prev, { id: nextId, ...form }]);
    setNextId(n => n + 1);
    setForm({ name:"", cat:"", maker:"", link:"" });
    notify("품목이 추가되었습니다.");
  };

  const deleteItem = (id) => {
    setDb(prev => prev.filter(x => x.id !== id));
    setSelected(prev => { const n = {...prev}; delete n[id]; return n; });
    notify("삭제되었습니다.");
  };

  const toggleSelect = (id) =>
    setSelected(prev => { const n = {...prev}; if (n[id]) delete n[id]; else n[id] = {qty:1}; return n; });

  const updateQty = (id, val) =>
    setSelected(prev => ({...prev, [id]: {qty: Math.max(1, parseInt(val)||1)}}));

  const getSelectedItems = () =>
    Object.keys(selected).map(id => {
      const item = db.find(x => x.id == id);
      return item ? {...item, qty: selected[id].qty} : null;
    }).filter(Boolean);

  const downloadCSV = () => {
    const items = getSelectedItems();
    if (!items.length) { notify("품목을 선택해주세요."); return; }
    const rows = [["카탈로그번호","품목명","제조사","수량","링크"], ...items.map(x => [x.cat, x.name, x.maker, x.qty, x.link])];
    const csv = rows.map(r => r.map(c => `"${String(c||"").replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF"+csv], {type:"text/csv;charset=utf-8;"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url;
    const today = new Date().toISOString().slice(0,10);
    a.download = `주문서_${today}.csv`; a.click(); URL.revokeObjectURL(url);
    const newEntry = { date: today, items: items.map(x => ({ name:x.name, cat:x.cat, maker:x.maker, link:x.link, qty:x.qty })) };
    setHistory(prev => [newEntry, ...prev].slice(0, 30));
    notify("CSV 다운로드 완료!");
  };

  const buildEmailBody = () => {
    const items = getSelectedItems();
    const today = new Date().toLocaleDateString("ko-KR");
    let body = `안녕하세요,\n\n아래 물품 주문 요청드립니다.\n\n주문일: ${today}\n\n${"─".repeat(28)}\n`;
    items.forEach((x, i) => {
      body += `${i+1}. ${x.name}\n   카탈로그번호: ${x.cat}\n`;
      if (x.maker) body += `   제조사: ${x.maker}\n`;
      body += `   수량: ${x.qty}\n`;
      if (x.link) body += `   링크: ${x.link}\n`;
      body += "\n";
    });
    return body + `${"─".repeat(28)}\n\n감사합니다.`;
  };

  const openGmailModal = () => {
    if (!getSelectedItems().length) { notify("품목을 선택해주세요."); return; }
    setGmailSubject(`[주문 요청] 연구실 물품 ${new Date().toLocaleDateString("ko-KR")}`);
    setGmailOpen(true);
  };

  const openGmailCompose = () => {
    const url = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(gmailTo)}&su=${encodeURIComponent(gmailSubject)}&body=${encodeURIComponent(buildEmailBody())}`;
    window.open(url, "_blank");
    setGmailOpen(false);
  };

  const reloadOrder = (idx) => {
    const h = history[idx];
    const newSel = {};
    h.items.forEach(hi => {
      const found = db.find(x => x.cat === hi.cat && x.name === hi.name);
      if (found) newSel[found.id] = { qty: hi.qty };
    });
    setSelected(newSel); setTab("order");
    notify(`${h.date} 주문 내역을 불러왔습니다.`);
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        if (!rows.length) { notify("데이터가 없습니다."); return; }
        const headers = Object.keys(rows[0]);
        const autoFind = (candidates) =>
          headers.find(h => candidates.some(c => h.replace(/\s/g,"").toLowerCase().includes(c))) || "";
        const mapping = {
          name:  autoFind(["품목명","name","item","품명","제품명"]),
          cat:   autoFind(["카탈로그","catalog","cat","번호","cat#","catno"]),
          maker: autoFind(["제조사","maker","manufacturer","brand","회사"]),
          link:  autoFind(["링크","link","url","홈페이지"]),
        };
        setMapModal({ headers, rows, mapping });
        e.target.value = "";
      } catch {
        notify("파일을 읽을 수 없습니다. 엑셀 또는 CSV 파일을 사용해주세요.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const applyMapping = () => {
    const { rows, mapping } = mapModal;
    let added = 0, skipped = 0;
    setDb(prev => {
      let updated = [...prev];
      let id = nextId;
      rows.forEach(row => {
        const name  = String(row[mapping.name]  ?? "").trim();
        const cat   = String(row[mapping.cat]   ?? "").trim();
        const maker = String(row[mapping.maker] ?? "").trim();
        const link  = String(row[mapping.link]  ?? "").trim();
        if (!name && !cat) { skipped++; return; }
        const dup = updated.find(x => x.cat === cat && x.name === name);
        if (dup) { skipped++; return; }
        updated.push({ id: id++, name, cat, maker, link });
        added++;
      });
      setNextId(id);
      return updated;
    });
    setMapModal(null);
    notify(`${added}개 추가됨${skipped ? `, ${skipped}개 중복/빈값 제외` : ""}`);
  };

  const filtered = db.filter(x =>
    !search || x.name.toLowerCase().includes(search.toLowerCase()) || x.cat.toLowerCase().includes(search.toLowerCase())
  );
  const selectedCount = Object.keys(selected).length;

  const S = {
    app: { fontFamily:"'Apple SD Gothic Neo','Malgun Gothic',sans-serif", maxWidth:660, margin:"0 auto", padding:"16px 16px 60px" },
    header: { padding:"14px 0 12px", display:"flex", alignItems:"center", gap:10, borderBottom:"1.5px solid #f0f0f0", marginBottom:18 },
    headerTitle: { fontSize:19, fontWeight:700, color:"#111" },
    headerSub: { fontSize:13, color:"#aaa", marginLeft:"auto" },
    tabs: { display:"flex", borderBottom:"1.5px solid #e5e7eb", marginBottom:18, gap:2 },
    tab: { padding:"8px 16px", fontSize:14, background:"none", border:"none", borderBottom:"2.5px solid transparent", marginBottom:-1.5, cursor:"pointer", color:"#999", fontWeight:500, borderRadius:"6px 6px 0 0" },
    tabActive: { color:"#1a1a1a", borderBottomColor:"#1a1a1a", fontWeight:700 },
    searchInput: { width:"100%", padding:"9px 14px", fontSize:14, border:"1.5px solid #e8e8e8", borderRadius:9, marginBottom:12, outline:"none", background:"#fafafa", boxSizing:"border-box" },
    itemList: { display:"flex", flexDirection:"column", gap:7 },
    itemRow: { display:"flex", alignItems:"center", padding:"11px 14px", border:"1.5px solid #ececec", borderRadius:11, background:"#fff", transition:"border-color 0.15s" },
    itemRowSel: { borderColor:"#2563eb", background:"#f0f6ff" },
    itemName: { fontSize:14, fontWeight:600, color:"#1a1a1a" },
    itemMeta: { fontSize:12, color:"#999", marginTop:2 },
    summaryBar: { display:"flex", alignItems:"center", justifyContent:"space-between", padding:"12px 14px", background:"#f7f7f8", borderRadius:11, marginTop:14, flexWrap:"wrap", gap:8 },
    card: { background:"#fff", border:"1.5px solid #ececec", borderRadius:13, padding:"18px 18px 14px", marginBottom:10 },
    cardTitle: { fontSize:15, fontWeight:700, marginBottom:14, color:"#111" },
    input: { padding:"8px 12px", fontSize:14, border:"1.5px solid #e0e0e0", borderRadius:8, outline:"none", width:"100%", boxSizing:"border-box", background:"#fafafa" },
    label: { display:"block", fontSize:12, color:"#888", marginBottom:4, fontWeight:500 },
    historyItem: { padding:"13px 16px", border:"1.5px solid #ececec", borderRadius:11, marginBottom:7, cursor:"pointer", background:"#fff" },
    tag: { fontSize:11, padding:"3px 9px", borderRadius:99, border:"1px solid #e5e5e5", color:"#555", background:"#f5f5f5" },
    empty: { textAlign:"center", padding:"40px 20px", color:"#bbb", fontSize:14 },
    btnPrimary: { padding:"8px 16px", fontSize:13, background:"#1a1a1a", color:"#fff", border:"none", borderRadius:8, cursor:"pointer", fontWeight:600 },
    btnOutline: { padding:"8px 14px", fontSize:13, background:"#fff", color:"#333", border:"1.5px solid #ccc", borderRadius:8, cursor:"pointer" },
    btnGhost: { padding:"8px 14px", fontSize:13, background:"transparent", color:"#777", border:"1.5px solid #e5e5e5", borderRadius:8, cursor:"pointer" },
    btnDanger: { padding:"5px 12px", fontSize:12, background:"#fff", color:"#e53e3e", border:"1px solid #fca5a5", borderRadius:7, cursor:"pointer" },
    qtyInput: { width:54, padding:"5px 8px", fontSize:13, border:"1.5px solid #ddd", borderRadius:7, textAlign:"center", outline:"none" },
    modalBg: { position:"fixed", inset:0, background:"rgba(0,0,0,0.32)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:100 },
    modal: { background:"#fff", borderRadius:15, padding:"26px 24px 22px", width:"92%", maxWidth:440, boxShadow:"0 8px 40px rgba(0,0,0,0.16)" },
    preview: { background:"#f6f6f6", borderRadius:9, padding:"12px 14px", fontSize:12, whiteSpace:"pre-wrap", lineHeight:1.75, maxHeight:160, overflowY:"auto", color:"#444" },
  };

  return (
    <div style={S.app}>
      <div style={S.header}>
        <span style={S.headerTitle}>🧪 연구실 주문 관리</span>
        <span style={S.headerSub}>{db.length}개 품목 등록됨</span>
      </div>

      <div style={S.tabs}>
        {[["order","🛒 주문하기"], ["db","📦 품목 DB"], ["history","📋 주문 이력"]].map(([id, label]) => (
          <button key={id} style={{...S.tab, ...(tab===id ? S.tabActive : {})}} onClick={() => setTab(id)}>{label}</button>
        ))}
      </div>

      {tab==="order" && (
        <div>
          <input style={S.searchInput} placeholder="🔍  품목명 또는 카탈로그번호 검색..." value={search} onChange={e => setSearch(e.target.value)} />
          {filtered.length === 0
            ? <div style={S.empty}>품목이 없습니다. 📦 품목 DB 탭에서 추가해주세요.</div>
            : <div style={S.itemList}>
                {filtered.map(item => (
                  <div key={item.id} style={{...S.itemRow, ...(selected[item.id] ? S.itemRowSel : {})}}>
                    <input type="checkbox" checked={!!selected[item.id]} onChange={() => toggleSelect(item.id)}
                      style={{marginRight:10, width:16, height:16, cursor:"pointer", flexShrink:0, accentColor:"#2563eb"}} />
                    <div style={{flex:1, minWidth:0}}>
                      <div style={S.itemName}>{item.name}</div>
                      <div style={S.itemMeta}>
                        {item.cat}{item.maker ? ` · ${item.maker}` : ""}
                        {item.link ? <a href={item.link} target="_blank" rel="noreferrer" style={{color:"#2563eb", marginLeft:6}}>링크 ↗</a> : ""}
                      </div>
                    </div>
                    {selected[item.id] && (
                      <div style={{display:"flex", alignItems:"center", gap:5}}>
                        <span style={{fontSize:12, color:"#999"}}>수량</span>
                        <input type="number" min="1" value={selected[item.id].qty}
                          onChange={e => updateQty(item.id, e.target.value)} style={S.qtyInput} />
                      </div>
                    )}
                  </div>
                ))}
              </div>
          }
          <div style={S.summaryBar}>
            <span style={{fontSize:13, color:"#666"}}>
              {selectedCount > 0 ? `✅ ${selectedCount}개 품목 선택됨` : "품목을 선택하세요"}
            </span>
            <div style={{display:"flex", gap:8}}>
              <button style={S.btnGhost} onClick={() => setSelected({})}>초기화</button>
              <button style={S.btnOutline} onClick={openGmailModal}>✉️ Gmail</button>
              <button style={S.btnPrimary} onClick={downloadCSV}>⬇️ CSV</button>
            </div>
          </div>
        </div>
      )}

      {tab==="db" && (
        <div>
          <div style={S.card}>
            <div style={S.cardTitle}>새 품목 추가</div>
            <div style={{display:"flex", gap:8, marginBottom:9, flexWrap:"wrap"}}>
              <input style={{...S.input, flex:"1 1 140px"}} placeholder="품목명 *" value={form.name}
                onChange={e => setForm(f => ({...f, name:e.target.value}))} />
              <input style={{...S.input, flex:"1 1 120px"}} placeholder="카탈로그번호 *" value={form.cat}
                onChange={e => setForm(f => ({...f, cat:e.target.value}))} />
            </div>
            <div style={{display:"flex", gap:8, marginBottom:13, flexWrap:"wrap"}}>
              <input style={{...S.input, flex:"1 1 120px"}} placeholder="제조사" value={form.maker}
                onChange={e => setForm(f => ({...f, maker:e.target.value}))} />
              <input style={{...S.input, flex:"2 1 180px"}} placeholder="품목 링크 (URL)" value={form.link}
                onChange={e => setForm(f => ({...f, link:e.target.value}))} />
            </div>
            <div style={{textAlign:"right"}}><button style={S.btnPrimary} onClick={addItem}>+ 추가</button></div>
          </div>

          <div style={{display:"flex", alignItems:"center", gap:10, margin:"6px 0 14px", padding:"11px 14px", background:"#f0f6ff", borderRadius:10, border:"1.5px dashed #93c5fd"}}>
            <span style={{fontSize:13, color:"#2563eb", fontWeight:600}}>📂 엑셀 업로드</span>
            <span style={{fontSize:12, color:"#888", flex:1}}>.xlsx / .xls / .csv · 중복 자동 제외</span>
            <label style={{...S.btnPrimary, padding:"6px 14px", cursor:"pointer"}}>
              파일 선택
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} style={{display:"none"}} />
            </label>
          </div>

          <div style={{fontSize:13, fontWeight:600, color:"#666", margin:"14px 0 8px"}}>등록된 품목 ({db.length})</div>
          {db.length === 0
            ? <div style={S.empty}>등록된 품목이 없습니다.</div>
            : <div style={S.itemList}>
                {db.map(item => (
                  <div key={item.id} style={S.itemRow}>
                    <div style={{flex:1, minWidth:0}}>
                      <div style={S.itemName}>{item.name}</div>
                      <div style={S.itemMeta}>
                        {item.cat}{item.maker ? ` · ${item.maker}` : ""}
                        {item.link ? <a href={item.link} target="_blank" rel="noreferrer" style={{color:"#2563eb", marginLeft:6}}>링크 ↗</a> : ""}
                      </div>
                    </div>
                    <button style={S.btnDanger} onClick={() => deleteItem(item.id)}>삭제</button>
                  </div>
                ))}
              </div>
          }
        </div>
      )}

      {tab==="history" && (
        <div>
          {history.length === 0
            ? <div style={S.empty}>아직 주문 이력이 없습니다.<br /><span style={{fontSize:12}}>CSV 다운로드 시 이력이 저장됩니다.</span></div>
            : history.map((h, idx) => (
                <div key={idx} style={S.historyItem} onClick={() => reloadOrder(idx)}>
                  <div style={{display:"flex", justifyContent:"space-between", alignItems:"center"}}>
                    <span style={{fontSize:13, color:"#555", fontWeight:600}}>📅 {h.date}</span>
                    <span style={{fontSize:12, color:"#2563eb"}}>재주문 →</span>
                  </div>
                  <div style={{marginTop:7, display:"flex", flexWrap:"wrap", gap:4}}>
                    {h.items.map((x, i) => <span key={i} style={S.tag}>{x.name} ×{x.qty}</span>)}
                  </div>
                </div>
              ))
          }
        </div>
      )}

      {gmailOpen && (
        <div style={S.modalBg} onClick={() => setGmailOpen(false)}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.cardTitle}>✉️ Gmail 초안 작성</div>
            <div style={{marginBottom:10}}>
              <label style={S.label}>받는 사람</label>
              <input style={S.input} placeholder="supplier@example.com" value={gmailTo} onChange={e => setGmailTo(e.target.value)} />
            </div>
            <div style={{marginBottom:10}}>
              <label style={S.label}>제목</label>
              <input style={S.input} value={gmailSubject} onChange={e => setGmailSubject(e.target.value)} />
            </div>
            <div style={{marginBottom:16}}>
              <label style={S.label}>본문 미리보기</label>
              <div style={S.preview}>{buildEmailBody()}</div>
            </div>
            <div style={{display:"flex", gap:8, justifyContent:"flex-end"}}>
              <button style={S.btnGhost} onClick={() => setGmailOpen(false)}>취소</button>
              <button style={S.btnPrimary} onClick={openGmailCompose}>Gmail에서 열기 ↗</button>
            </div>
          </div>
        </div>
      )}

      {mapModal && (
        <div style={S.modalBg} onClick={() => setMapModal(null)}>
          <div style={{...S.modal, maxWidth:480}} onClick={e => e.stopPropagation()}>
            <div style={S.cardTitle}>📋 컬럼 매핑 확인</div>
            <div style={{fontSize:12, color:"#888", marginBottom:14}}>엑셀의 어떤 컬럼이 어떤 필드인지 선택해주세요.</div>
            {[
              { key:"name",  label:"품목명 *" },
              { key:"cat",   label:"카탈로그번호 *" },
              { key:"maker", label:"제조사" },
              { key:"link",  label:"링크" },
            ].map(({ key, label }) => (
              <div key={key} style={{display:"flex", alignItems:"center", gap:10, marginBottom:10}}>
                <span style={{fontSize:13, fontWeight:600, color:"#444", width:110, flexShrink:0}}>{label}</span>
                <select value={mapModal.mapping[key]}
                  onChange={e => setMapModal(m => ({...m, mapping:{...m.mapping, [key]:e.target.value}}))}
                  style={{...S.input, background:"#fff", flex:1}}>
                  <option value="">(매핑 안 함)</option>
                  {mapModal.headers.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
            <div style={{margin:"14px 0 6px", fontSize:12, color:"#888", fontWeight:600}}>미리보기 (최대 3행)</div>
            <div style={{overflowX:"auto", marginBottom:16}}>
              <table style={{width:"100%", fontSize:11, borderCollapse:"collapse"}}>
                <thead>
                  <tr style={{background:"#f5f5f5"}}>
                    {["품목명","카탈로그번호","제조사","링크"].map(h => (
                      <th key={h} style={{padding:"5px 8px", border:"1px solid #eee", textAlign:"left", color:"#555"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {mapModal.rows.slice(0,3).map((row, i) => (
                    <tr key={i}>
                      {["name","cat","maker","link"].map(k => (
                        <td key={k} style={{padding:"5px 8px", border:"1px solid #eee", color:"#333", maxWidth:100, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>
                          {mapModal.mapping[k] ? String(row[mapModal.mapping[k]] ?? "") : <span style={{color:"#ccc"}}>-</span>}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{display:"flex", gap:8, justifyContent:"flex-end"}}>
              <button style={S.btnGhost} onClick={() => setMapModal(null)}>취소</button>
              <button style={S.btnPrimary} onClick={applyMapping}>가져오기 ({mapModal.rows.length}행)</button>
            </div>
          </div>
        </div>
      )}

      <Notification msg={notif} />
    </div>
  );
}