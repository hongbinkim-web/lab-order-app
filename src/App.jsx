import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── 주문 관리 ───────────────────────────────────────────
const ORDER_SAMPLE = [
  { id: 1, name: "RPMI 1640 배지", cat: "R8758", maker: "Sigma-Aldrich", link: "https://www.sigmaaldrich.com" },
  { id: 2, name: "FBS (Fetal Bovine Serum)", cat: "16000044", maker: "Gibco", link: "" },
  { id: 3, name: "Trypan Blue Solution", cat: "T8154", maker: "Sigma-Aldrich", link: "" },
];

// ─── Cell Stock ───────────────────────────────────────────
const ROWS = ["A","B","C","D","E","F","G","H","I"];
const COLS = [1,2,3,4,5,6,7,8,9];

const STOCK_SAMPLE = [
  {
    id: 1, name: "F1-A-3",
    cells: {
      "A3":"Virus Flag.DLL4.28z 11/20/26 HN","A4":"Virus Flag.SBT01M.28z 11/7/26 HN","A5":"Virus Flag.SBT01G.28z 11/7/26 HN",
      "B4":"Jurkat (NFAT_GFP) 10M p.7 10/24/25 HB","B5":"Jurkat (NFAT_GFP) 10M p.7 10/24/25 HB","B6":"Jurkat (NFAT_GFP) 10M p.7 10/24/25 HB","B7":"Jurkat (NFAT_GFP) 10M p.7 10/24/25 HB",
      "C4":"Jurkat (ATCC) 10M p.5 11/11/25 HB","C5":"Jurkat (ATCC) 10M p.5 11/11/25 HB","C6":"Jurkat (ATCC) 10M p.5 11/11/25 HB","C7":"Jurkat (ATCC) 10M p.5 11/11/25 HB",
      "D2":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D3":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D4":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D5":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D6":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D7":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D8":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB","D9":"Flag.DLL4.28z-Jurkat (Sort X) 10M p.5 12/2/25 HB",
      "F1":"Flag.SBT01G.28z-Jurkat (Sort O) 5M 12/12/25 HB","F2":"Flag.SBT01G.28z-Jurkat (Sort O) 5M 12/12/25 HB","F3":"Flag.SBT01G.28z-Jurkat (Sort O) 5M 12/12/25 HB","F4":"Flag.SBT01G.28z-Jurkat (Sort O) 5M 12/12/25 HB","F5":"Flag.SBT01G.28z-Jurkat (Sort O) 5M 12/12/25 HB",
      "G1":"Flag.SBT01M.28z-Jurkat (Sort O) 15M 12/17/25 HB","G2":"Flag.SBT01M.28z-Jurkat (Sort O) 15M 12/17/25 HB","G3":"Flag.SBT01M.28z-Jurkat (Sort O) 15M 12/17/25 HB",
      "I3":"Flag.SBT01M.28z-Pan T (LRS 24-2) 10M 12/29/25 HB","I4":"Flag.SBT01G.28z-Pan T (LRS 24-2) 10M 12/29/25 HB","I5":"Flag.DLL4.28z-Pan T (LRS 24-2) 10M 12/29/25 HB","I6":"NTD-Pan T (LRS 24-2) 10M 12/29/25 HB",
    }
  }
];

function getCellColor(content) {
  if (!content) return null;
  const s = content.toLowerCase();
  if (s.includes("jurkat")) return "#3b82f6";
  if (s.includes("nalm6")) return "#8b5cf6";
  if (s.includes("293t")) return "#10b981";
  if (s.includes("virus")) return "#f59e0b";
  if (s.includes("pan t")) return "#ef4444";
  return "#6b7280";
}

function Notification({ msg }) {
  if (!msg) return null;
  return <div style={{position:"fixed",bottom:24,right:24,background:"#1a1a1a",color:"#fff",padding:"10px 18px",borderRadius:9,fontSize:13,zIndex:999,boxShadow:"0 4px 16px rgba(0,0,0,0.18)"}}>{msg}</div>;
}

// ─── 주문 관리 컴포넌트 ───────────────────────────────────
function OrderApp({ notify }) {
  const [db, setDb] = useState(ORDER_SAMPLE);
  const [history, setHistory] = useState([]);
  const [selected, setSelected] = useState({});
  const [search, setSearch] = useState("");
  const [subTab, setSubTab] = useState("order");
  const [form, setForm] = useState({ name:"", cat:"", maker:"", link:"" });
  const [gmailOpen, setGmailOpen] = useState(false);
  const [gmailTo, setGmailTo] = useState("");
  const [gmailSubject, setGmailSubject] = useState("");
  const [nextId, setNextId] = useState(100);
  const [mapModal, setMapModal] = useState(null);

  const addItem = () => {
    if (!form.name.trim() || !form.cat.trim()) { notify("품목명과 카탈로그번호는 필수입니다."); return; }
    setDb(prev => [...prev, { id: nextId, ...form }]);
    setNextId(n => n+1); setForm({ name:"", cat:"", maker:"", link:"" });
    notify("품목이 추가되었습니다.");
  };
  const deleteItem = (id) => { setDb(prev => prev.filter(x => x.id !== id)); setSelected(prev => { const n={...prev}; delete n[id]; return n; }); notify("삭제되었습니다."); };
  const toggleSelect = (id) => setSelected(prev => { const n={...prev}; if(n[id]) delete n[id]; else n[id]={qty:1}; return n; });
  const updateQty = (id, val) => setSelected(prev => ({...prev, [id]:{qty:Math.max(1,parseInt(val)||1)}}));
  const getSelectedItems = () => Object.keys(selected).map(id => { const item=db.find(x=>x.id==id); return item?{...item,qty:selected[id].qty}:null; }).filter(Boolean);

  const downloadCSV = () => {
    const items = getSelectedItems();
    if (!items.length) { notify("품목을 선택해주세요."); return; }
    const rows = [["카탈로그번호","품목명","제조사","수량","링크"], ...items.map(x=>[x.cat,x.name,x.maker,x.qty,x.link])];
    const csv = rows.map(r=>r.map(c=>`"${String(c||"").replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
    const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href=url;
    const today = new Date().toISOString().slice(0,10);
    a.download=`주문서_${today}.csv`; a.click(); URL.revokeObjectURL(url);
    setHistory(prev=>[{date:today,items:items.map(x=>({name:x.name,cat:x.cat,maker:x.maker,link:x.link,qty:x.qty}))}, ...prev].slice(0,30));
    notify("CSV 다운로드 완료!");
  };

  const buildEmailBody = () => {
    const items = getSelectedItems(); const today = new Date().toLocaleDateString("ko-KR");
    let body=`안녕하세요,\n\n아래 물품 주문 요청드립니다.\n\n주문일: ${today}\n\n${"─".repeat(28)}\n`;
    items.forEach((x,i)=>{ body+=`${i+1}. ${x.name}\n   카탈로그번호: ${x.cat}\n`; if(x.maker) body+=`   제조사: ${x.maker}\n`; body+=`   수량: ${x.qty}\n`; if(x.link) body+=`   링크: ${x.link}\n`; body+="\n"; });
    return body+`${"─".repeat(28)}\n\n감사합니다.`;
  };

  const openGmailModal = () => {
    if (!getSelectedItems().length) { notify("품목을 선택해주세요."); return; }
    setGmailSubject(`[주문 요청] 연구실 물품 ${new Date().toLocaleDateString("ko-KR")}`); setGmailOpen(true);
  };
  const openGmailCompose = () => { window.open(`https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(gmailTo)}&su=${encodeURIComponent(gmailSubject)}&body=${encodeURIComponent(buildEmailBody())}`,"_blank"); setGmailOpen(false); };

  const reloadOrder = (idx) => {
    const h=history[idx]; const newSel={};
    h.items.forEach(hi=>{ const found=db.find(x=>x.cat===hi.cat&&x.name===hi.name); if(found) newSel[found.id]={qty:hi.qty}; });
    setSelected(newSel); setSubTab("order"); notify(`${h.date} 주문 내역을 불러왔습니다.`);
  };

  const handleFileUpload = (e) => {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(evt)=>{
      try {
        const wb=XLSX.read(evt.target.result,{type:"binary"}); const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false}); if(!rows.length){notify("데이터가 없습니다.");return;}
        const headers=Object.keys(rows[0]);
        const autoFind=(candidates)=>headers.find(h=>candidates.some(c=>h.replace(/\s/g,"").toLowerCase().includes(c)))||"";
        const mapping={name:autoFind(["품목명","name","item","품명","제품명"]),cat:autoFind(["카탈로그","catalog","cat","번호"]),maker:autoFind(["제조사","maker","manufacturer"]),link:autoFind(["링크","link","url"])};
        setMapModal({headers,rows,mapping}); e.target.value="";
      } catch { notify("파일을 읽을 수 없습니다."); }
    };
    reader.readAsBinaryString(file);
  };

  const applyMapping = () => {
    const {rows,mapping}=mapModal; let added=0,skipped=0;
    setDb(prev=>{ let updated=[...prev]; let id=nextId;
      rows.forEach(row=>{ const name=String(row[mapping.name]??"").trim(); const cat=String(row[mapping.cat]??"").trim(); const maker=String(row[mapping.maker]??"").trim(); const link=String(row[mapping.link]??"").trim();
        if(!name&&!cat){skipped++;return;} const dup=updated.find(x=>x.cat===cat&&x.name===name); if(dup){skipped++;return;} updated.push({id:id++,name,cat,maker,link}); added++;
      }); setNextId(id); return updated;
    }); setMapModal(null); notify(`${added}개 추가됨${skipped?`, ${skipped}개 중복/빈값 제외`:""}`)
  };

  const filtered = db.filter(x=>!search||x.name.toLowerCase().includes(search.toLowerCase())||x.cat.toLowerCase().includes(search.toLowerCase()));
  const selectedCount = Object.keys(selected).length;

  const S = {
    searchInput:{width:"100%",padding:"9px 14px",fontSize:14,border:"1.5px solid #e8e8e8",borderRadius:9,marginBottom:12,outline:"none",background:"#fafafa",boxSizing:"border-box"},
    itemRow:{display:"flex",alignItems:"center",padding:"11px 14px",border:"1.5px solid #ececec",borderRadius:11,background:"#fff",marginBottom:7},
    itemRowSel:{borderColor:"#2563eb",background:"#f0f6ff"},
    itemName:{fontSize:14,fontWeight:600,color:"#1a1a1a"},
    itemMeta:{fontSize:12,color:"#999",marginTop:2},
    summaryBar:{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",background:"#f7f7f8",borderRadius:11,marginTop:14,flexWrap:"wrap",gap:8},
    card:{background:"#fff",border:"1.5px solid #ececec",borderRadius:13,padding:"18px 18px 14px",marginBottom:10},
    cardTitle:{fontSize:15,fontWeight:700,marginBottom:14,color:"#111"},
    input:{padding:"8px 12px",fontSize:14,border:"1.5px solid #e0e0e0",borderRadius:8,outline:"none",width:"100%",boxSizing:"border-box",background:"#fafafa"},
    label:{display:"block",fontSize:12,color:"#888",marginBottom:4,fontWeight:500},
    tag:{fontSize:11,padding:"3px 9px",borderRadius:99,border:"1px solid #e5e5e5",color:"#555",background:"#f5f5f5"},
    empty:{textAlign:"center",padding:"40px 20px",color:"#bbb",fontSize:14},
    btnPrimary:{padding:"8px 16px",fontSize:13,background:"#1a1a1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600},
    btnOutline:{padding:"8px 14px",fontSize:13,background:"#fff",color:"#333",border:"1.5px solid #ccc",borderRadius:8,cursor:"pointer"},
    btnGhost:{padding:"8px 14px",fontSize:13,background:"transparent",color:"#777",border:"1.5px solid #e5e5e5",borderRadius:8,cursor:"pointer"},
    btnDanger:{padding:"5px 12px",fontSize:12,background:"#fff",color:"#e53e3e",border:"1px solid #fca5a5",borderRadius:7,cursor:"pointer"},
    qtyInput:{width:54,padding:"5px 8px",fontSize:13,border:"1.5px solid #ddd",borderRadius:7,textAlign:"center",outline:"none"},
    modalBg:{position:"fixed",inset:0,background:"rgba(0,0,0,0.32)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
    modal:{background:"#fff",borderRadius:15,padding:"26px 24px 22px",width:"92%",maxWidth:440,boxShadow:"0 8px 40px rgba(0,0,0,0.16)"},
    preview:{background:"#f6f6f6",borderRadius:9,padding:"12px 14px",fontSize:12,whiteSpace:"pre-wrap",lineHeight:1.75,maxHeight:160,overflowY:"auto",color:"#444"},
    subTabs:{display:"flex",borderBottom:"1.5px solid #e5e7eb",marginBottom:16,gap:2},
    subTab:{padding:"7px 14px",fontSize:13,background:"none",border:"none",borderBottom:"2px solid transparent",marginBottom:-1.5,cursor:"pointer",color:"#999",fontWeight:500},
    subTabActive:{color:"#1a1a1a",borderBottomColor:"#1a1a1a",fontWeight:700},
  };

  return (
    <div>
      <div style={S.subTabs}>
        {[["order","🛒 주문하기"],["db","📦 품목 DB"],["history","📋 주문 이력"]].map(([id,label])=>(
          <button key={id} style={{...S.subTab,...(subTab===id?S.subTabActive:{})}} onClick={()=>setSubTab(id)}>{label}</button>
        ))}
      </div>

      {subTab==="order"&&(
        <div>
          <input style={S.searchInput} placeholder="🔍  품목명 또는 카탈로그번호 검색..." value={search} onChange={e=>setSearch(e.target.value)}/>
          {filtered.length===0?<div style={S.empty}>품목이 없습니다. 📦 품목 DB 탭에서 추가해주세요.</div>:
            filtered.map(item=>(
              <div key={item.id} style={{...S.itemRow,...(selected[item.id]?S.itemRowSel:{})}}>
                <input type="checkbox" checked={!!selected[item.id]} onChange={()=>toggleSelect(item.id)} style={{marginRight:10,width:16,height:16,cursor:"pointer",flexShrink:0,accentColor:"#2563eb"}}/>
                <div style={{flex:1,minWidth:0}}>
                  <div style={S.itemName}>{item.name}</div>
                  <div style={S.itemMeta}>{item.cat}{item.maker?` · ${item.maker}`:""}{item.link?<a href={item.link} target="_blank" rel="noreferrer" style={{color:"#2563eb",marginLeft:6}}>링크 ↗</a>:""}</div>
                </div>
                {selected[item.id]&&<div style={{display:"flex",alignItems:"center",gap:5}}><span style={{fontSize:12,color:"#999"}}>수량</span><input type="number" min="1" value={selected[item.id].qty} onChange={e=>updateQty(item.id,e.target.value)} style={S.qtyInput}/></div>}
              </div>
            ))
          }
          <div style={S.summaryBar}>
            <span style={{fontSize:13,color:"#666"}}>{selectedCount>0?`✅ ${selectedCount}개 품목 선택됨`:"품목을 선택하세요"}</span>
            <div style={{display:"flex",gap:8}}>
              <button style={S.btnGhost} onClick={()=>setSelected({})}>초기화</button>
              <button style={S.btnOutline} onClick={openGmailModal}>✉️ Gmail</button>
              <button style={S.btnPrimary} onClick={downloadCSV}>⬇️ CSV</button>
            </div>
          </div>
        </div>
      )}

      {subTab==="db"&&(
        <div>
          <div style={S.card}>
            <div style={S.cardTitle}>새 품목 추가</div>
            <div style={{display:"flex",gap:8,marginBottom:9,flexWrap:"wrap"}}>
              <input style={{...S.input,flex:"1 1 140px"}} placeholder="품목명 *" value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/>
              <input style={{...S.input,flex:"1 1 120px"}} placeholder="카탈로그번호 *" value={form.cat} onChange={e=>setForm(f=>({...f,cat:e.target.value}))}/>
            </div>
            <div style={{display:"flex",gap:8,marginBottom:13,flexWrap:"wrap"}}>
              <input style={{...S.input,flex:"1 1 120px"}} placeholder="제조사" value={form.maker} onChange={e=>setForm(f=>({...f,maker:e.target.value}))}/>
              <input style={{...S.input,flex:"2 1 180px"}} placeholder="품목 링크 (URL)" value={form.link} onChange={e=>setForm(f=>({...f,link:e.target.value}))}/>
            </div>
            <div style={{textAlign:"right"}}><button style={S.btnPrimary} onClick={addItem}>+ 추가</button></div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10,margin:"6px 0 14px",padding:"11px 14px",background:"#f0f6ff",borderRadius:10,border:"1.5px dashed #93c5fd"}}>
            <span style={{fontSize:13,color:"#2563eb",fontWeight:600}}>📂 엑셀 업로드</span>
            <span style={{fontSize:12,color:"#888",flex:1}}>.xlsx / .xls / .csv · 중복 자동 제외</span>
            <label style={{...S.btnPrimary,padding:"6px 14px",cursor:"pointer"}}>파일 선택<input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} style={{display:"none"}}/></label>
          </div>
          <div style={{fontSize:13,fontWeight:600,color:"#666",margin:"14px 0 8px"}}>등록된 품목 ({db.length})</div>
          {db.length===0?<div style={S.empty}>등록된 품목이 없습니다.</div>:
            db.map(item=>(
              <div key={item.id} style={S.itemRow}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={S.itemName}>{item.name}</div>
                  <div style={S.itemMeta}>{item.cat}{item.maker?` · ${item.maker}`:""}{item.link?<a href={item.link} target="_blank" rel="noreferrer" style={{color:"#2563eb",marginLeft:6}}>링크 ↗</a>:""}</div>
                </div>
                <button style={S.btnDanger} onClick={()=>deleteItem(item.id)}>삭제</button>
              </div>
            ))
          }
        </div>
      )}

      {subTab==="history"&&(
        <div>
          {history.length===0?<div style={S.empty}>아직 주문 이력이 없습니다.<br/><span style={{fontSize:12}}>CSV 다운로드 시 이력이 저장됩니다.</span></div>:
            history.map((h,idx)=>(
              <div key={idx} style={{padding:"13px 16px",border:"1.5px solid #ececec",borderRadius:11,marginBottom:7,cursor:"pointer",background:"#fff"}} onClick={()=>reloadOrder(idx)}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:13,color:"#555",fontWeight:600}}>📅 {h.date}</span>
                  <span style={{fontSize:12,color:"#2563eb"}}>재주문 →</span>
                </div>
                <div style={{marginTop:7,display:"flex",flexWrap:"wrap",gap:4}}>
                  {h.items.map((x,i)=><span key={i} style={S.tag}>{x.name} ×{x.qty}</span>)}
                </div>
              </div>
            ))
          }
        </div>
      )}

      {gmailOpen&&(
        <div style={S.modalBg} onClick={()=>setGmailOpen(false)}>
          <div style={S.modal} onClick={e=>e.stopPropagation()}>
            <div style={S.cardTitle}>✉️ Gmail 초안 작성</div>
            <div style={{marginBottom:10}}><label style={S.label}>받는 사람</label><input style={S.input} placeholder="supplier@example.com" value={gmailTo} onChange={e=>setGmailTo(e.target.value)}/></div>
            <div style={{marginBottom:10}}><label style={S.label}>제목</label><input style={S.input} value={gmailSubject} onChange={e=>setGmailSubject(e.target.value)}/></div>
            <div style={{marginBottom:16}}><label style={S.label}>본문 미리보기</label><div style={S.preview}>{buildEmailBody()}</div></div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button style={S.btnGhost} onClick={()=>setGmailOpen(false)}>취소</button>
              <button style={S.btnPrimary} onClick={openGmailCompose}>Gmail에서 열기 ↗</button>
            </div>
          </div>
        </div>
      )}

      {mapModal&&(
        <div style={S.modalBg} onClick={()=>setMapModal(null)}>
          <div style={{...S.modal,maxWidth:480}} onClick={e=>e.stopPropagation()}>
            <div style={S.cardTitle}>📋 컬럼 매핑 확인</div>
            <div style={{fontSize:12,color:"#888",marginBottom:14}}>엑셀의 어떤 컬럼이 어떤 필드인지 선택해주세요.</div>
            {[{key:"name",label:"품목명 *"},{key:"cat",label:"카탈로그번호 *"},{key:"maker",label:"제조사"},{key:"link",label:"링크"}].map(({key,label})=>(
              <div key={key} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                <span style={{fontSize:13,fontWeight:600,color:"#444",width:110,flexShrink:0}}>{label}</span>
                <select value={mapModal.mapping[key]} onChange={e=>setMapModal(m=>({...m,mapping:{...m.mapping,[key]:e.target.value}}))} style={{...S.input,background:"#fff",flex:1}}>
                  <option value="">(매핑 안 함)</option>
                  {mapModal.headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
            <div style={{margin:"14px 0 6px",fontSize:12,color:"#888",fontWeight:600}}>미리보기 (최대 3행)</div>
            <div style={{overflowX:"auto",marginBottom:16}}>
              <table style={{width:"100%",fontSize:11,borderCollapse:"collapse"}}>
                <thead><tr style={{background:"#f5f5f5"}}>{["품목명","카탈로그번호","제조사","링크"].map(h=><th key={h} style={{padding:"5px 8px",border:"1px solid #eee",textAlign:"left",color:"#555"}}>{h}</th>)}</tr></thead>
                <tbody>{mapModal.rows.slice(0,3).map((row,i)=><tr key={i}>{["name","cat","maker","link"].map(k=><td key={k} style={{padding:"5px 8px",border:"1px solid #eee",color:"#333",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{mapModal.mapping[k]?String(row[mapModal.mapping[k]]??""):<span style={{color:"#ccc"}}>-</span>}</td>)}</tr>)}</tbody>
              </table>
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button style={S.btnGhost} onClick={()=>setMapModal(null)}>취소</button>
              <button style={S.btnPrimary} onClick={applyMapping}>가져오기 ({mapModal.rows.length}행)</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Cell Stock 컴포넌트 ──────────────────────────────────
function StockApp({ notify }) {
  const [boxes, setBoxes] = useState(STOCK_SAMPLE);
  const [selectedBox, setSelectedBox] = useState(null);
  const [selectedCell, setSelectedCell] = useState(null);
  const [search, setSearch] = useState("");
  const [subTab, setSubTab] = useState("map");
  const [editModal, setEditModal] = useState(null);
  const [addBoxModal, setAddBoxModal] = useState(false);
  const [newBoxName, setNewBoxName] = useState("");
  const [nextId, setNextId] = useState(100);
  const [importModal, setImportModal] = useState(null);

  const searchResults = search.trim().length > 1
    ? boxes.flatMap(box => Object.entries(box.cells).filter(([,v])=>v.toLowerCase().includes(search.toLowerCase())).map(([cellKey,content])=>({boxName:box.name,boxId:box.id,cellKey,content})))
    : [];

  const updateCell = (boxId, cellKey, content) => {
    setBoxes(prev=>prev.map(b=>b.id===boxId?{...b,cells:content?{...b.cells,[cellKey]:content}:Object.fromEntries(Object.entries(b.cells).filter(([k])=>k!==cellKey))}:b));
  };

  const addBox = () => {
    if (!newBoxName.trim()) { notify("박스 이름을 입력해주세요."); return; }
    setBoxes(prev=>[...prev,{id:nextId,name:newBoxName.trim(),cells:{}}]);
    setNextId(n=>n+1); setNewBoxName(""); setAddBoxModal(false); notify("박스가 추가되었습니다.");
  };

  const deleteBox = (id) => { setBoxes(prev=>prev.filter(b=>b.id!==id)); if(selectedBox?.id===id) setSelectedBox(null); notify("박스가 삭제되었습니다."); };

  const handleFileUpload = (e) => {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(evt)=>{
      try {
        const wb=XLSX.read(evt.target.result,{type:"binary"});
        let parsedBoxes=[];

        // 모든 시트 순회
        wb.SheetNames.forEach(sheetName=>{
          const ws=wb.Sheets[sheetName];
          const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});

          // 각 행을 순회하며 박스 찾기
          for(let i=0;i<raw.length;i++){
            const row=raw[i];

            // "Box Name" 포함한 열 찾기
            const boxNameCols=[];
            row.forEach((cell,ci)=>{ if(String(cell).toLowerCase().includes("box name")) boxNameCols.push(ci); });
            if(!boxNameCols.length) continue;

            // 각 Box Name 열에 대해 처리
            boxNameCols.forEach(bnCol=>{
              // #R.n 헤더 행 찾기 (Box Name 행 바로 다음)
              const headerRow=raw[i+1]||[];

              // 컬럼 1~9 위치 찾기: #R.n 헤더에서 숫자 1~9 찾기
              const colPositions={}; // colPositions[1] = 실제 열 인덱스
              headerRow.forEach((cell,ci)=>{
                const n=parseInt(String(cell).trim());
                if(n>=1&&n<=9) colPositions[n]=ci;
              });

              // 행 레이블(A~I) 열 위치 찾기
              // Box Name 열 기준으로 왼쪽이나 같은 열에서 A~I 레이블 찾기
              let rowLabelCol=bnCol; // 기본값

              // 데이터 행들 파싱 (i+2부터 최대 11행)
              const cells={};
              let boxName=`${sheetName}_Box${parsedBoxes.length+1}`;

              for(let r=i+2;r<Math.min(i+13,raw.length);r++){
                const dataRow=raw[r];

                // 행 레이블(A~I) 찾기: 해당 행에서 A~I 값 찾기
                let letter=null;
                for(let ci=0;ci<Math.min(dataRow.length, bnCol+2);ci++){
                  const v=String(dataRow[ci]||"").trim();
                  if(ROWS.includes(v)){ letter=v; rowLabelCol=ci; break; }
                }
                if(!letter) continue;

                // 박스 이름: 행 레이블 왼쪽 셀에서 찾기
                if(!boxName.includes("_Box")){} // 이미 설정됨
                const leftCell=String(dataRow[rowLabelCol-1]||"").trim();
                if(leftCell && !ROWS.includes(leftCell) && !leftCell.includes("Box") && leftCell.length<20){
                  boxName=leftCell;
                }

                // 1~9 컬럼 데이터 읽기
                if(Object.keys(colPositions).length>0){
                  // 헤더에서 찾은 컬럼 위치 사용
                  for(let c=1;c<=9;c++){
                    if(colPositions[c]!==undefined){
                      const val=String(dataRow[colPositions[c]]||"").trim();
                      if(val) cells[`${letter}${c}`]=val;
                    }
                  }
                } else {
                  // 폴백: bnCol+1 부터 9개
                  for(let c=1;c<=9;c++){
                    const val=String(dataRow[bnCol+c]||"").trim();
                    if(val) cells[`${letter}${c}`]=val;
                  }
                }
              }

              if(Object.keys(cells).length>0){
                parsedBoxes.push({id:Date.now()+parsedBoxes.length+Math.random(),name:boxName,cells});
              }
            });
          }
        });

        if(!parsedBoxes.length){notify("박스 데이터를 찾을 수 없습니다. 파일 구조를 확인해주세요.");return;}
        setImportModal({boxes:parsedBoxes}); e.target.value="";
      } catch(err) { notify("파일을 읽을 수 없습니다: "+err.message); }
    };
    reader.readAsBinaryString(file);
  };

  const applyImport = (mode) => {
    if(mode==="replace") setBoxes(importModal.boxes);
    else setBoxes(prev=>[...prev,...importModal.boxes]);
    setImportModal(null); notify(`${importModal.boxes.length}개 박스 불러왔습니다.`);
  };

  const box = selectedBox ? boxes.find(b=>b.id===selectedBox.id) : null;
  const legendItems=[{label:"Jurkat",color:"#3b82f6"},{label:"Nalm6",color:"#8b5cf6"},{label:"293T",color:"#10b981"},{label:"Virus",color:"#f59e0b"},{label:"Pan T",color:"#ef4444"},{label:"기타",color:"#6b7280"}];

  const S = {
    subTabs:{display:"flex",borderBottom:"1.5px solid #e5e7eb",marginBottom:16,gap:2},
    subTab:{padding:"7px 14px",fontSize:13,background:"none",border:"none",borderBottom:"2px solid transparent",marginBottom:-1.5,cursor:"pointer",color:"#999",fontWeight:500},
    subTabActive:{color:"#1a1a1a",borderBottomColor:"#1a1a1a",fontWeight:700},
    boxGrid:{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(130px, 1fr))",gap:8,marginBottom:18},
    boxCard:{padding:"10px 12px",border:"1.5px solid #ececec",borderRadius:10,cursor:"pointer",background:"#fff",fontSize:13,fontWeight:600},
    boxCardActive:{borderColor:"#2563eb",background:"#f0f6ff"},
    th:{padding:"4px 6px",background:"#f5f5f5",border:"1px solid #e0e0e0",color:"#888",fontWeight:600,minWidth:70},
    td:{padding:0,border:"1px solid #e8e8e8",width:72,height:52,cursor:"pointer",position:"relative"},
    cellInner:{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",padding:"2px 3px",fontSize:10,lineHeight:1.3,textAlign:"center",overflow:"hidden"},
    searchInput:{width:"100%",padding:"10px 14px",fontSize:14,border:"1.5px solid #e8e8e8",borderRadius:9,outline:"none",background:"#fafafa",boxSizing:"border-box",marginBottom:14},
    resultItem:{padding:"12px 16px",border:"1.5px solid #ececec",borderRadius:11,marginBottom:7,background:"#fff",cursor:"pointer"},
    badge:{display:"inline-block",padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:700,color:"#fff",marginBottom:4},
    btnPrimary:{padding:"8px 16px",fontSize:13,background:"#1a1a1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600},
    btnOutline:{padding:"8px 14px",fontSize:13,background:"#fff",color:"#333",border:"1.5px solid #ccc",borderRadius:8,cursor:"pointer"},
    btnGhost:{padding:"8px 14px",fontSize:13,background:"transparent",color:"#777",border:"1.5px solid #e5e5e5",borderRadius:8,cursor:"pointer"},
    btnDanger:{padding:"6px 12px",fontSize:12,background:"#fff",color:"#e53e3e",border:"1px solid #fca5a5",borderRadius:7,cursor:"pointer"},
    modalBg:{position:"fixed",inset:0,background:"rgba(0,0,0,0.32)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
    modal:{background:"#fff",borderRadius:15,padding:"26px 24px 22px",width:"92%",maxWidth:440,boxShadow:"0 8px 40px rgba(0,0,0,0.16)"},
    input:{padding:"8px 12px",fontSize:14,border:"1.5px solid #e0e0e0",borderRadius:8,outline:"none",width:"100%",boxSizing:"border-box",background:"#fafafa"},
    label:{display:"block",fontSize:12,color:"#888",marginBottom:4,fontWeight:500},
    empty:{textAlign:"center",padding:"40px 20px",color:"#bbb",fontSize:14},
  };

  return (
    <div>
      <div style={S.subTabs}>
        {[["map","🗺️ 박스 맵"],["search","🔍 검색"]].map(([id,label])=>(
          <button key={id} style={{...S.subTab,...(subTab===id?S.subTabActive:{})}} onClick={()=>setSubTab(id)}>{label}</button>
        ))}
      </div>

      {subTab==="map"&&(
        <div>
          <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
            <button style={S.btnPrimary} onClick={()=>setAddBoxModal(true)}>+ 박스 추가</button>
            <label style={{...S.btnOutline,cursor:"pointer"}}>📂 엑셀 업로드<input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{display:"none"}}/></label>
            {selectedBox&&<button style={S.btnDanger} onClick={()=>deleteBox(selectedBox.id)}>🗑️ {boxes.find(b=>b.id===selectedBox.id)?.name} 삭제</button>}
          </div>
          <div style={S.boxGrid}>
            {boxes.map(b=>{
              const filled=Object.keys(b.cells).length;
              return <div key={b.id} style={{...S.boxCard,...(selectedBox?.id===b.id?S.boxCardActive:{})}} onClick={()=>setSelectedBox(selectedBox?.id===b.id?null:b)}>
                <div>{b.name}</div><div style={{fontSize:11,color:"#999",marginTop:3}}>{filled}/81 칸 사용</div>
              </div>;
            })}
          </div>
          {box&&(
            <div>
              <div style={{fontSize:15,fontWeight:700,marginBottom:8,color:"#111"}}>
                📦 {box.name} <span style={{fontSize:12,color:"#888",fontWeight:400}}>칸 클릭 → 추가/수정/삭제</span>
              </div>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,margin:"8px 0 12px"}}>
                {legendItems.map(l=><div key={l.label} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#555"}}><div style={{width:10,height:10,borderRadius:3,background:l.color}}/>{l.label}</div>)}
              </div>
              <div style={{overflowX:"auto"}}>
                <table style={{borderCollapse:"collapse",fontSize:11}}>
                  <thead><tr><th style={S.th}></th>{COLS.map(c=><th key={c} style={S.th}>{c}</th>)}</tr></thead>
                  <tbody>
                    {ROWS.map(row=>(
                      <tr key={row}>
                        <td style={{...S.th,background:"#f5f5f5"}}>{row}</td>
                        {COLS.map(col=>{
                          const key=`${row}${col}`; const content=box.cells[key]||""; const color=getCellColor(content); const isSel=selectedCell===key;
                          return <td key={col} style={{...S.td,background:color?`${color}18`:"#fff",outline:isSel?`2px solid ${color||"#2563eb"}`:"none"}}
                            onClick={()=>{setSelectedCell(key);setEditModal({boxId:box.id,cellKey:key,content});}}>
                            <div style={{...S.cellInner,color:color||"#bbb"}}>
                              {content?<span title={content}>{content.length>30?content.slice(0,28)+"…":content}</span>:<span style={{fontSize:16,color:"#e0e0e0"}}>+</span>}
                            </div>
                          </td>;
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          {!box&&<div style={S.empty}>박스를 선택하면 그리드가 표시됩니다</div>}
        </div>
      )}

      {subTab==="search"&&(
        <div>
          <input style={S.searchInput} placeholder="🔍  세포주 이름 검색... (예: Jurkat, DLL4, Pan T)" value={search} onChange={e=>setSearch(e.target.value)} autoFocus/>
          {search.trim().length>1&&searchResults.length===0&&<div style={S.empty}>검색 결과가 없습니다.</div>}
          {searchResults.map((r,i)=>{
            const color=getCellColor(r.content);
            return <div key={i} style={S.resultItem} onClick={()=>{setSelectedBox({id:r.boxId});setSubTab("map");}}>
              <div style={{...S.badge,background:color||"#6b7280"}}>{r.boxName} — {r.cellKey}</div>
              <div style={{fontSize:13,color:"#333"}}>{r.content}</div>
            </div>;
          })}
          {search.trim().length<=1&&<div style={S.empty}>2글자 이상 입력하세요</div>}
        </div>
      )}

      {editModal&&(
        <div style={S.modalBg} onClick={()=>{setEditModal(null);setSelectedCell(null);}}>
          <div style={S.modal} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:16,color:"#111"}}>📍 {boxes.find(b=>b.id===editModal.boxId)?.name} — {editModal.cellKey}</div>
            <div style={{marginBottom:16}}>
              <label style={S.label}>내용 (세포주명 · 세포수 · 날짜 · 담당자)</label>
              <textarea value={editModal.content} onChange={e=>setEditModal(m=>({...m,content:e.target.value}))} style={{...S.input,height:80,resize:"vertical",fontFamily:"inherit"}} placeholder="예: Jurkat (ATCC) 10M p.5 11/11/25 HB" autoFocus/>
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"space-between"}}>
              <button style={S.btnDanger} onClick={()=>{updateCell(editModal.boxId,editModal.cellKey,"");setEditModal(null);setSelectedCell(null);notify("삭제되었습니다.");}}>🗑️ 삭제</button>
              <div style={{display:"flex",gap:8}}>
                <button style={S.btnGhost} onClick={()=>{setEditModal(null);setSelectedCell(null);}}>취소</button>
                <button style={S.btnPrimary} onClick={()=>{updateCell(editModal.boxId,editModal.cellKey,editModal.content);setEditModal(null);setSelectedCell(null);notify("저장되었습니다.");}}>저장</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {addBoxModal&&(
        <div style={S.modalBg} onClick={()=>setAddBoxModal(false)}>
          <div style={S.modal} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:16}}>📦 새 박스 추가</div>
            <div style={{marginBottom:16}}><label style={S.label}>박스 이름</label><input style={S.input} placeholder="예: F1-A-3" value={newBoxName} onChange={e=>setNewBoxName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addBox()} autoFocus/></div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button style={S.btnGhost} onClick={()=>setAddBoxModal(false)}>취소</button>
              <button style={S.btnPrimary} onClick={addBox}>추가</button>
            </div>
          </div>
        </div>
      )}

      {importModal&&(
        <div style={S.modalBg}>
          <div style={S.modal}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:10}}>📂 엑셀 불러오기</div>
            <div style={{fontSize:13,color:"#555",marginBottom:16}}><b>{importModal.boxes.length}개</b> 박스를 찾았습니다.<br/>{importModal.boxes.map(b=>b.name).join(", ")}</div>
            <div style={{fontSize:13,color:"#888",marginBottom:20}}>기존 데이터를 어떻게 할까요?</div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end",flexWrap:"wrap"}}>
              <button style={S.btnGhost} onClick={()=>setImportModal(null)}>취소</button>
              <button style={S.btnOutline} onClick={()=>applyImport("merge")}>기존 유지 + 추가</button>
              <button style={S.btnPrimary} onClick={()=>applyImport("replace")}>전체 교체</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── 메인 앱 ─────────────────────────────────────────────
export default function App() {
  const [mainTab, setMainTab] = useState("order");
  const [notif, setNotif] = useState("");
  const notify = useCallback((msg)=>{ setNotif(msg); setTimeout(()=>setNotif(""),2400); },[]);

  const S = {
    app:{fontFamily:"'Apple SD Gothic Neo','Malgun Gothic',sans-serif",maxWidth:820,margin:"0 auto",padding:"16px 16px 60px"},
    header:{padding:"14px 0 14px",display:"flex",alignItems:"center",gap:10,borderBottom:"1.5px solid #f0f0f0",marginBottom:0},
    headerTitle:{fontSize:19,fontWeight:700,color:"#111"},
    mainTabs:{display:"flex",background:"#f5f5f5",borderRadius:12,padding:4,gap:4,margin:"16px 0 20px"},
    mainTab:{flex:1,padding:"10px",fontSize:14,background:"transparent",border:"none",borderRadius:9,cursor:"pointer",color:"#888",fontWeight:600,transition:"all 0.15s"},
    mainTabActive:{background:"#fff",color:"#111",boxShadow:"0 1px 4px rgba(0,0,0,0.10)"},
  };

  return (
    <div style={S.app}>
      <div style={S.header}>
        <span style={S.headerTitle}>🔬 연구실 통합 관리</span>
      </div>
      <div style={S.mainTabs}>
        <button style={{...S.mainTab,...(mainTab==="order"?S.mainTabActive:{})}} onClick={()=>setMainTab("order")}>🧪 주문 관리</button>
        <button style={{...S.mainTab,...(mainTab==="stock"?S.mainTabActive:{})}} onClick={()=>setMainTab("stock")}>🧊 Cell Stock Map</button>
      </div>
      {mainTab==="order" && <OrderApp notify={notify}/>}
      {mainTab==="stock" && <StockApp notify={notify}/>}
      <Notification msg={notif}/>
    </div>
  );
}