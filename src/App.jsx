import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const ROWS = ["A","B","C","D","E","F","G","H","I"];
const COLS = [1,2,3,4,5,6,7,8,9];
const ROW_SET = new Set(ROWS);

// ── 색상 ──────────────────────────────────────────────────
function getCellColor(content) {
  if (!content) return null;
  const s = content.toLowerCase();
  if (s.includes("jurkat")) return "#3b82f6";
  if (s.includes("nalm6") || s.includes("nalm-6")) return "#8b5cf6";
  if (s.includes("293t") || s.includes("expi293")) return "#10b981";
  if (s.includes("virus") || s.includes("lenti")) return "#f59e0b";
  if (s.includes("pan t") || s.includes("pbmc") || s.includes("lrs")) return "#ef4444";
  if (s.includes("mb3") || s.includes("u87")) return "#ec4899";
  return "#6b7280";
}

const LEGEND = [
  {label:"Jurkat", color:"#3b82f6"},
  {label:"Nalm6", color:"#8b5cf6"},
  {label:"293T/Expi", color:"#10b981"},
  {label:"Virus/Lenti", color:"#f59e0b"},
  {label:"T/PBMC", color:"#ef4444"},
  {label:"기타", color:"#6b7280"},
];

// ── 샘플 데이터 ───────────────────────────────────────────
        const INIT_STORAGE = {
  deepfreezer: [
    {
      id:1, name:"F1-A-3",
      cells:{
        "B4":"Jurkat (NFAT_GFP) 10M p.7 10/24/25 HB","B5":"Jurkat (NFAT_GFP) 10M p.7 10/24/25 HB",
        "C4":"Jurkat (ATCC) 10M p.5 11/11/25 HB","C5":"Jurkat (ATCC) 10M p.5 11/11/25 HB",
        "D2":"Flag.DLL4.28z-Jurkat 10M 12/2/25 HB","D3":"Flag.DLL4.28z-Jurkat 10M 12/2/25 HB",
        "I3":"Flag.SBT01M.28z-Pan T 10M 12/29/25 HB","I4":"Flag.DLL4.28z-Pan T 10M 12/29/25 HB",
      }
    }
  ],
  ln2: {
    // sheetName → { rackId → { boxLetter → { cells } } }
  }
};

function Notification({ msg }) {
  if (!msg) return null;
  return <div style={{position:"fixed",bottom:24,right:24,background:"#1a1a1a",color:"#fff",padding:"10px 18px",borderRadius:9,fontSize:13,zIndex:999,boxShadow:"0 4px 16px rgba(0,0,0,0.18)"}}>{msg}</div>;
}

// ── 그리드 컴포넌트 ───────────────────────────────────────
function CellGrid({ cells, onCellClick }) {
  return (
    <div style={{overflowX:"auto"}}>
      <table style={{borderCollapse:"collapse",fontSize:11}}>
        <thead>
          <tr>
            <th style={{padding:"4px 8px",background:"#f5f5f5",border:"1px solid #e0e0e0",color:"#888",fontWeight:600,minWidth:28}}></th>
            {COLS.map(c=><th key={c} style={{padding:"4px 6px",background:"#f5f5f5",border:"1px solid #e0e0e0",color:"#888",fontWeight:600,minWidth:72,textAlign:"center"}}>{c}</th>)}
          </tr>
        </thead>
        <tbody>
          {ROWS.map(row=>(
            <tr key={row}>
              <td style={{padding:"4px 8px",background:"#f5f5f5",border:"1px solid #e0e0e0",color:"#888",fontWeight:600,textAlign:"center"}}>{row}</td>
              {COLS.map(col=>{
                const key=`${row}${col}`;
                const content=cells[key]||"";
                const color=getCellColor(content);
                return (
                  <td key={col} onClick={()=>onCellClick(key,content)}
                    style={{padding:0,border:"1px solid #e8e8e8",width:72,height:52,cursor:"pointer",background:color?`${color}18`:"#fff",transition:"background 0.1s"}}>
                    <div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",padding:"2px 3px",fontSize:10,lineHeight:1.3,textAlign:"center",overflow:"hidden",color:color||"#bbb"}}>
                      {content
                        ? <span title={content}>{content.length>32?content.slice(0,30)+"…":content}</span>
                        : <span style={{fontSize:16,color:"#e8e8e8"}}>+</span>
                      }
                    </div>
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── 셀 편집 모달 ──────────────────────────────────────────
function EditModal({ cellKey, content, onSave, onDelete, onClose }) {
  const [val, setVal] = useState(content);
  const S = {
    modalBg:{position:"fixed",inset:0,background:"rgba(0,0,0,0.32)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200},
    modal:{background:"#fff",borderRadius:15,padding:"24px",width:"92%",maxWidth:420,boxShadow:"0 8px 40px rgba(0,0,0,0.16)"},
    input:{padding:"8px 12px",fontSize:14,border:"1.5px solid #e0e0e0",borderRadius:8,outline:"none",width:"100%",boxSizing:"border-box",background:"#fafafa",height:80,resize:"vertical",fontFamily:"inherit"},
    btnP:{padding:"8px 16px",fontSize:13,background:"#1a1a1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600},
    btnG:{padding:"8px 14px",fontSize:13,background:"transparent",color:"#777",border:"1.5px solid #e5e5e5",borderRadius:8,cursor:"pointer"},
    btnD:{padding:"6px 12px",fontSize:12,background:"#fff",color:"#e53e3e",border:"1px solid #fca5a5",borderRadius:7,cursor:"pointer"},
  };
  return (
    <div style={S.modalBg} onClick={onClose}>
      <div style={S.modal} onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:15,fontWeight:700,marginBottom:14,color:"#111"}}>📍 {cellKey} 편집</div>
        <textarea value={val} onChange={e=>setVal(e.target.value)} style={S.input}
          placeholder="예: Jurkat (ATCC) 10M p.5 11/11/25 HB" autoFocus/>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:14}}>
          <button style={S.btnD} onClick={()=>onDelete(cellKey)}>🗑️ 삭제</button>
          <div style={{display:"flex",gap:8}}>
            <button style={S.btnG} onClick={onClose}>취소</button>
            <button style={S.btnP} onClick={()=>onSave(cellKey,val)}>저장</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Deepfreezer 탭 ────────────────────────────────────────
function DeepfreezerTab({ notify, boxes, setBoxes }) {
  const [selBox, setSelBox] = useState(null);
  const [editModal, setEditModal] = useState(null);
  const [addModal, setAddModal] = useState(false);
  const [newName, setNewName] = useState("");
  const [search, setSearch] = useState("");
  const [subTab, setSubTab] = useState("map");
  const [nextId, setNextId] = useState(50);
  const [importModal, setImportModal] = useState(null);

  const box = selBox ? boxes.find(b=>b.id===selBox) : null;

  const updateCell = (boxId, key, val) => {
    setBoxes(prev=>prev.map(b=>b.id===boxId
      ? {...b, cells: val ? {...b.cells,[key]:val} : Object.fromEntries(Object.entries(b.cells).filter(([k])=>k!==key))}
      : b));
  };

  const searchResults = search.trim().length > 1
    ? boxes.flatMap(b=>Object.entries(b.cells).filter(([,v])=>v.toLowerCase().includes(search.toLowerCase())).map(([k,v])=>({boxName:b.name,boxId:b.id,key:k,content:v})))
    : [];

  const handleFileUpload = (e) => {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(evt)=>{
      try {
        const wb=XLSX.read(evt.target.result,{type:"binary"});
        const ws=wb.Sheets["Deepfreezer"]||wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});

        // 결과: boxName → cells
        const boxMap={};
        const getBox=(name)=>{ if(!boxMap[name]) boxMap[name]={cells:{}}; return boxMap[name]; };

        let i=0;
        while(i<raw.length){
          const row=raw[i]||[];
          const rowStr=row.map(c=>String(c??"").trim());

          // ── 패턴 1: "Box Name : xxx" 행 ──
          const boxNameCol=rowStr.findIndex(v=>v.toLowerCase().includes("box name"));
          if(boxNameCol>=0){
            // 같은 행에 여러 Box Name이 있을 수 있음
            const boxNameGroups=[];
            rowStr.forEach((v,ci)=>{
              if(!v.toLowerCase().includes("box name")) return;
              // Box Name 값: 같은 셀에 ":"뒤에 있거나 옆 셀
              let bname=v.replace(/box name\s*:?\s*/i,"").trim();
              if(!bname){ for(let k=ci+1;k<=ci+2;k++){ const nv=String((row[k]??"")).trim(); if(nv){bname=nv;break;} } }
              boxNameGroups.push({ci, bname:bname||"Unknown"});
            });

            // 다음 행: #R.n 헤더
            const headerRow=raw[i+1]||[];
            const rnGroups=[];
            headerRow.forEach((cell,ci)=>{
              if(String(cell??"").trim()!=="#R.n") return;
              const colPos={};
              for(let k=ci+1;k<headerRow.length&&k<ci+15;k++){
                const n=parseInt(String(headerRow[k]??"").trim());
                if(!isNaN(n)&&n>=1&&n<=9) colPos[n]=k;
              }
              if(Object.keys(colPos).length>=3) rnGroups.push({ci,colPos});
            });

            // 각 #R.n 그룹과 가장 가까운 Box Name 매핑
            const groupBoxNames=rnGroups.map(({ci:rci})=>{
              let closest=null, minDist=9999;
              boxNameGroups.forEach(({ci:bci,bname})=>{
                const d=Math.abs(rci-bci);
                if(d<minDist){minDist=d;closest=bname;}
              });
              return closest||"Unknown";
            });

            // 데이터 행 파싱 (i+2부터)
            const dataStartRow=i+2;
            // 박스 위치명(F1-A-x) 추적
            const groupBoxPos=rnGroups.map(()=>null);

            for(let r=dataStartRow;r<raw.length;r++){
              const dr=raw[r]||[];
              if(dr.every(c=>String(c??"").trim()==="")) break;
              if(dr.some(c=>String(c??"").toLowerCase().includes("box name"))) break;
              if(dr.some(c=>String(c??"").trim()==="#R.n")) break;

              rnGroups.forEach(({ci:hci,colPos},gi)=>{
                // 박스 위치명 업데이트
                for(let k=Math.max(0,hci-3);k<=hci;k++){
                  const v=String(dr[k]??"").trim();
                  if(v&&v.match(/^F\d+-[A-Z]-\d+$/i)){groupBoxPos[gi]=v;break;}
                }
                // 행 레이블
                let letter=null;
                for(let k=Math.max(0,hci-1);k<=hci+2&&k<dr.length;k++){
                  if(ROW_SET.has(String(dr[k]??"").trim())){letter=String(dr[k]??"").trim();break;}
                }
                if(!letter) return;

                const bname=groupBoxNames[gi];
                const box=getBox(bname);
                for(let c=1;c<=9;c++){
                  const ci2=colPos[c]; if(ci2==null) continue;
                  const val=String(dr[ci2]??"").trim();
                  if(val) box.cells[`${letter}${c}`]=val;
                }
              });
            }
            i+=2; continue;
          }

          // ── 패턴 2: "#R.n" 직접 헤더 ──
          const isRnRow=rowStr.some(v=>v==="#R.n");
          if(isRnRow){
            const rnGroups=[];
            rowStr.forEach((v,ci)=>{
              if(v!=="#R.n") return;
              const colPos={};
              for(let k=ci+1;k<row.length&&k<ci+15;k++){
                const n=parseInt(String(row[k]??"").trim());
                if(!isNaN(n)&&n>=1&&n<=9) colPos[n]=k;
              }
              if(Object.keys(colPos).length>=3) rnGroups.push({ci,colPos});
            });

            const groupBoxNames=rnGroups.map(()=>null);

            for(let r=i+1;r<raw.length;r++){
              const dr=raw[r]||[];
              if(dr.every(c=>String(c??"").trim()==="")) break;
              if(dr.some(c=>String(c??"").toLowerCase().includes("box name"))) break;
              if(dr.some(c=>String(c??"").trim()==="#R.n")) break;

              rnGroups.forEach(({ci:hci,colPos},gi)=>{
                // 박스명: hci 왼쪽에서 F1-A-x 패턴 찾기
                for(let k=Math.max(0,hci-3);k<=hci;k++){
                  const v=String(dr[k]??"").trim();
                  if(v&&v.length>2&&!ROW_SET.has(v)&&!/^\d+$/.test(v)){
                    groupBoxNames[gi]=v; break;
                  }
                }
                let letter=null;
                for(let k=Math.max(0,hci-1);k<=hci+2&&k<dr.length;k++){
                  if(ROW_SET.has(String(dr[k]??"").trim())){letter=String(dr[k]??"").trim();break;}
                }
                if(!letter||!groupBoxNames[gi]) return;

                const box=getBox(groupBoxNames[gi]);
                for(let c=1;c<=9;c++){
                  const ci2=colPos[c]; if(ci2==null) continue;
                  const val=String(dr[ci2]??"").trim();
                  if(val) box.cells[`${letter}${c}`]=val;
                }
              });
            }
          }
          i++;
        }

        const parsed=Object.entries(boxMap).map(([name,{cells}],idx)=>({
          id:Date.now()+idx, name, cells
        }));

        if(!parsed.length){notify("박스 데이터를 찾을 수 없습니다.");return;}
        setImportModal({boxes:parsed}); e.target.value="";
      } catch(err){notify("파일 읽기 오류: "+err.message);}
    };
    reader.readAsBinaryString(file);
  };

  const S = {
    subTabs:{display:"flex",borderBottom:"1.5px solid #e5e7eb",marginBottom:14,gap:2},
    subTab:{padding:"7px 14px",fontSize:13,background:"none",border:"none",borderBottom:"2px solid transparent",marginBottom:-1.5,cursor:"pointer",color:"#999",fontWeight:500},
    subTabActive:{color:"#1a1a1a",borderBottomColor:"#1a1a1a",fontWeight:700},
    boxGrid:{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:8,marginBottom:16},
    boxCard:{padding:"10px 12px",border:"1.5px solid #ececec",borderRadius:10,cursor:"pointer",background:"#fff",fontSize:13,fontWeight:600},
    boxCardActive:{borderColor:"#2563eb",background:"#f0f6ff"},
    btnP:{padding:"8px 16px",fontSize:13,background:"#1a1a1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600},
    btnO:{padding:"8px 14px",fontSize:13,background:"#fff",color:"#333",border:"1.5px solid #ccc",borderRadius:8,cursor:"pointer"},
    btnG:{padding:"8px 14px",fontSize:13,background:"transparent",color:"#777",border:"1.5px solid #e5e5e5",borderRadius:8,cursor:"pointer"},
    btnD:{padding:"6px 12px",fontSize:12,background:"#fff",color:"#e53e3e",border:"1px solid #fca5a5",borderRadius:7,cursor:"pointer"},
    input:{padding:"8px 12px",fontSize:14,border:"1.5px solid #e0e0e0",borderRadius:8,outline:"none",width:"100%",boxSizing:"border-box",background:"#fafafa"},
    searchInput:{width:"100%",padding:"10px 14px",fontSize:14,border:"1.5px solid #e8e8e8",borderRadius:9,outline:"none",background:"#fafafa",boxSizing:"border-box",marginBottom:14},
    modalBg:{position:"fixed",inset:0,background:"rgba(0,0,0,0.32)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
    modal:{background:"#fff",borderRadius:15,padding:"24px",width:"92%",maxWidth:440,boxShadow:"0 8px 40px rgba(0,0,0,0.16)"},
    label:{display:"block",fontSize:12,color:"#888",marginBottom:4,fontWeight:500},
    empty:{textAlign:"center",padding:"36px",color:"#bbb",fontSize:14},
    resultItem:{padding:"12px 16px",border:"1.5px solid #ececec",borderRadius:11,marginBottom:7,background:"#fff",cursor:"pointer"},
    badge:{display:"inline-block",padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:700,color:"#fff",marginBottom:4},
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
          <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
            <button style={S.btnP} onClick={()=>setAddModal(true)}>+ 박스 추가</button>
            <label style={{...S.btnO,cursor:"pointer"}}>📂 엑셀 업로드<input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{display:"none"}}/></label>
            {selBox&&<button style={S.btnD} onClick={()=>{setBoxes(p=>p.filter(b=>b.id!==selBox));setSelBox(null);notify("삭제됨");}}>🗑️ 박스 삭제</button>}
          </div>
          <div style={S.boxGrid}>
            {boxes.map(b=>(
              <div key={b.id} style={{...S.boxCard,...(selBox===b.id?S.boxCardActive:{})}} onClick={()=>setSelBox(selBox===b.id?null:b.id)}>
                <div>{b.name}</div>
                <div style={{fontSize:11,color:"#999",marginTop:3}}>{Object.keys(b.cells).length}/81칸</div>
              </div>
            ))}
            {boxes.length===0&&<div style={S.empty}>박스가 없습니다</div>}
          </div>
          {box&&(
            <div>
              <div style={{fontSize:14,fontWeight:700,marginBottom:8}}>📦 {box.name}</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
                {LEGEND.map(l=><div key={l.label} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#555"}}><div style={{width:10,height:10,borderRadius:3,background:l.color}}/>{l.label}</div>)}
              </div>
              <CellGrid cells={box.cells} onCellClick={(key,content)=>setEditModal({boxId:box.id,cellKey:key,content})}/>
            </div>
          )}
          {!box&&boxes.length>0&&<div style={S.empty}>박스를 선택하세요</div>}
        </div>
      )}

      {subTab==="search"&&(
        <div>
          <input style={S.searchInput} placeholder="🔍 세포주 이름 검색..." value={search} onChange={e=>setSearch(e.target.value)} autoFocus/>
          {search.trim().length>1&&searchResults.length===0&&<div style={S.empty}>검색 결과 없음</div>}
          {searchResults.map((r,i)=>{
            const color=getCellColor(r.content);
            return <div key={i} style={S.resultItem} onClick={()=>{setSelBox(r.boxId);setSubTab("map");}}>
              <div style={{...S.badge,background:color||"#6b7280"}}>{r.boxName} — {r.key}</div>
              <div style={{fontSize:13,color:"#333",marginTop:2}}>{r.content}</div>
            </div>;
          })}
          {search.trim().length<=1&&<div style={S.empty}>2글자 이상 입력하세요</div>}
        </div>
      )}

      {editModal&&(
        <EditModal cellKey={editModal.cellKey} content={editModal.content}
          onSave={(key,val)=>{updateCell(editModal.boxId,key,val);setEditModal(null);notify("저장됨");}}
          onDelete={(key)=>{updateCell(editModal.boxId,key,"");setEditModal(null);notify("삭제됨");}}
          onClose={()=>setEditModal(null)}/>
      )}

      {addModal&&(
        <div style={S.modalBg} onClick={()=>setAddModal(false)}>
          <div style={S.modal} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>📦 새 박스 추가</div>
            <label style={S.label}>박스 이름</label>
            <input style={S.input} placeholder="예: F1-A-3" value={newName} onChange={e=>setNewName(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){if(!newName.trim()){notify("이름 필수");return;}setBoxes(p=>[...p,{id:nextId,name:newName.trim(),cells:{}}]);setNextId(n=>n+1);setNewName("");setAddModal(false);notify("추가됨");}}} autoFocus/>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:14}}>
              <button style={S.btnG} onClick={()=>setAddModal(false)}>취소</button>
              <button style={S.btnP} onClick={()=>{if(!newName.trim()){notify("이름 필수");return;}setBoxes(p=>[...p,{id:nextId,name:newName.trim(),cells:{}}]);setNextId(n=>n+1);setNewName("");setAddModal(false);notify("추가됨");}}>추가</button>
            </div>
          </div>
        </div>
      )}

      {importModal&&(
        <div style={S.modalBg}>
          <div style={S.modal}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:10}}>📂 엑셀 불러오기</div>
            <div style={{fontSize:13,color:"#555",marginBottom:6}}><b>{importModal.boxes.length}개</b> 박스 발견:</div>
            <div style={{fontSize:12,color:"#888",marginBottom:16,maxHeight:80,overflowY:"auto"}}>{importModal.boxes.map(b=>b.name).join(", ")}</div>
            <div style={{fontSize:13,color:"#666",marginBottom:16}}>기존 데이터를 어떻게 할까요?</div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end",flexWrap:"wrap"}}>
              <button style={S.btnG} onClick={()=>setImportModal(null)}>취소</button>
              <button style={S.btnO} onClick={()=>{setBoxes(p=>[...p,...importModal.boxes]);setImportModal(null);notify(`${importModal.boxes.length}개 박스 추가됨`);}}>기존 유지 + 추가</button>
              <button style={S.btnP} onClick={()=>{setBoxes(importModal.boxes);setImportModal(null);notify(`${importModal.boxes.length}개 박스 불러옴`);}}>전체 교체</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── LN2 탭 ────────────────────────────────────────────────
function LN2Tab({ notify, ln2Data, setLn2Data }) {
  const [selSheet, setSelSheet] = useState(null);
  const [selRack, setSelRack] = useState(null);
  const [selBox, setSelBox] = useState(null);
  const [editModal, setEditModal] = useState(null);
  const [search, setSearch] = useState("");
  const [subTab, setSubTab] = useState("map");
  const [importModal, setImportModal] = useState(null);

  const sheets = Object.keys(ln2Data);
  const racks = selSheet ? Object.keys(ln2Data[selSheet]||{}).sort((a,b)=>parseInt(a)-parseInt(b)) : [];
  const boxes = (selSheet&&selRack) ? Object.keys(ln2Data[selSheet][selRack]||{}).sort() : [];
  const cells = (selSheet&&selRack&&selBox) ? (ln2Data[selSheet]?.[selRack]?.[selBox]?.cells||{}) : null;

  const updateCell = (key, val) => {
    setLn2Data(prev=>{
      const updated=JSON.parse(JSON.stringify(prev));
      if(!updated[selSheet]?.[selRack]?.[selBox]) return prev;
      if(val) updated[selSheet][selRack][selBox].cells[key]=val;
      else delete updated[selSheet][selRack][selBox].cells[key];
      return updated;
    });
  };

  const searchResults = search.trim().length > 1
    ? Object.entries(ln2Data).flatMap(([sheet,racks])=>
        Object.entries(racks).flatMap(([rack,boxes])=>
          Object.entries(boxes).flatMap(([box,{cells}])=>
            Object.entries(cells).filter(([,v])=>v.toLowerCase().includes(search.toLowerCase()))
              .map(([k,v])=>({sheet,rack,box,key:k,content:v}))
          )
        )
      )
    : [];

  const handleFileUpload = (e) => {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(evt)=>{
      try {
        const wb=XLSX.read(evt.target.result,{type:"binary"});
        const newData={};

        wb.SheetNames.forEach(sheetName=>{
          if(sheetName.toLowerCase()==="deepfreezer") return;
          const ws=wb.Sheets[sheetName];
          const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
          const sheetData={};

          for(let i=0;i<raw.length;i++){
            const row=raw[i]||[];

            // 헤더행 감지: #숫자 패턴이 있는 행
            // 구조: [#1] [A] [] [1] [2]...[9] [] [#2] [A] [] [1]...[9]
            const groups=[];
            for(let ci=0;ci<row.length;ci++){
              const v=String(row[ci]??"").trim();
              if(!/^#\d+$/.test(v)) continue;
              const rack=v.replace("#","");

              // 박스 레터: #숫자 바로 다음 셀 (ROW_SET 포함 모든 알파벳 허용)
              let box="A";
              for(let k=ci+1;k<=ci+3&&k<row.length;k++){
                const bv=String(row[k]??"").trim();
                if(/^[A-Z]$/.test(bv)){box=bv;break;}
              }

              // 1~9 컬럼 위치: ci 이후에서 찾기
              const colPos={};
              for(let k=ci+1;k<row.length&&k<ci+15;k++){
                const n=parseInt(String(row[k]??"").trim());
                if(!isNaN(n)&&n>=1&&n<=9) colPos[n]=k;
              }
              if(Object.keys(colPos).length>=5){
                groups.push({rack, box, colPos, headerCol:ci});
              }
            }
            if(groups.length===0) continue;

            // 데이터 행 파싱
            for(let r=i+1;r<raw.length;r++){
              const dr=raw[r]||[];
              if(dr.every(c=>String(c??"").trim()==="")) break;
              if(dr.some(c=>/^#\d+$/.test(String(c??"").trim()))) break;

              groups.forEach(({rack,box,colPos,headerCol})=>{
                // 행 레이블: headerCol 기준 ±2 범위에서 A~I 찾기
                let letter=null;
                for(let k=Math.max(0,headerCol-1);k<=headerCol+3&&k<dr.length;k++){
                  const v=String(dr[k]??"").trim();
                  if(ROW_SET.has(v)){letter=v;break;}
                }
                if(!letter) return;

                if(!sheetData[rack]) sheetData[rack]={};
                if(!sheetData[rack][box]) sheetData[rack][box]={cells:{}};

                for(let c=1;c<=9;c++){
                  const ci=colPos[c];
                  if(ci==null) continue;
                  const val=String(dr[ci]??"").trim();
                  if(val) sheetData[rack][box].cells[`${letter}${c}`]=val;
                }
              });
            }
          }

          if(Object.keys(sheetData).length>0) newData[sheetName]=sheetData;
        });

        if(!Object.keys(newData).length){notify("LN2 데이터를 찾을 수 없습니다.");return;}
        setImportModal({data:newData});
        e.target.value="";
      } catch(err){notify("파일 읽기 오류: "+err.message);}
    };
    reader.readAsBinaryString(file);
  };

  const S = {
    subTabs:{display:"flex",borderBottom:"1.5px solid #e5e7eb",marginBottom:14,gap:2},
    subTab:{padding:"7px 14px",fontSize:13,background:"none",border:"none",borderBottom:"2px solid transparent",marginBottom:-1.5,cursor:"pointer",color:"#999",fontWeight:500},
    subTabActive:{color:"#1a1a1a",borderBottomColor:"#1a1a1a",fontWeight:700},
    btnP:{padding:"8px 16px",fontSize:13,background:"#1a1a1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600},
    btnO:{padding:"8px 14px",fontSize:13,background:"#fff",color:"#333",border:"1.5px solid #ccc",borderRadius:8,cursor:"pointer"},
    btnG:{padding:"8px 14px",fontSize:13,background:"transparent",color:"#777",border:"1.5px solid #e5e5e5",borderRadius:8,cursor:"pointer"},
    modalBg:{position:"fixed",inset:0,background:"rgba(0,0,0,0.32)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
    modal:{background:"#fff",borderRadius:15,padding:"24px",width:"92%",maxWidth:440,boxShadow:"0 8px 40px rgba(0,0,0,0.16)"},
    chip:(active,color="#2563eb")=>({display:"inline-flex",alignItems:"center",padding:"6px 14px",borderRadius:99,fontSize:13,fontWeight:600,cursor:"pointer",border:`1.5px solid ${active?color:"#e0e0e0"}`,background:active?`${color}18`:"#fff",color:active?color:"#777",margin:"0 4px 6px 0"}),
    searchInput:{width:"100%",padding:"10px 14px",fontSize:14,border:"1.5px solid #e8e8e8",borderRadius:9,outline:"none",background:"#fafafa",boxSizing:"border-box",marginBottom:14},
    resultItem:{padding:"12px 16px",border:"1.5px solid #ececec",borderRadius:11,marginBottom:7,background:"#fff",cursor:"pointer"},
    badge:{display:"inline-block",padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:700,color:"#fff",marginBottom:4},
    empty:{textAlign:"center",padding:"36px",color:"#bbb",fontSize:14},
  };

  const rackColors=["#2563eb","#7c3aed","#059669","#d97706","#dc2626","#0891b2"];

  return (
    <div>
      <div style={S.subTabs}>
        {[["map","🗺️ 랙/박스 맵"],["search","🔍 검색"]].map(([id,label])=>(
          <button key={id} style={{...S.subTab,...(subTab===id?S.subTabActive:{})}} onClick={()=>setSubTab(id)}>{label}</button>
        ))}
      </div>

      {subTab==="map"&&(
        <div>
          <div style={{marginBottom:14}}>
            <label style={{...S.btnO,cursor:"pointer",display:"inline-block"}}>📂 엑셀 업로드<input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{display:"none"}}/></label>
          </div>

          {sheets.length===0&&(
            <div style={S.empty}>엑셀 파일을 업로드해서 LN2 데이터를 불러오세요</div>
          )}

          {/* 시트 선택 */}
          {sheets.length>0&&(
            <div style={{marginBottom:16}}>
              <div style={{fontSize:12,color:"#888",fontWeight:600,marginBottom:8}}>📁 저장소</div>
              <div>{sheets.map(s=><span key={s} style={S.chip(selSheet===s)} onClick={()=>{setSelSheet(s);setSelRack(null);setSelBox(null);}}>{s}</span>)}</div>
            </div>
          )}

          {/* 랙 선택 */}
          {selSheet&&racks.length>0&&(
            <div style={{marginBottom:16}}>
              <div style={{fontSize:12,color:"#888",fontWeight:600,marginBottom:8}}>🗂️ 랙</div>
              <div>{racks.map((r,ri)=><span key={r} style={S.chip(selRack===r, rackColors[ri%rackColors.length])} onClick={()=>{setSelRack(r);setSelBox(null);}}>랙 #{r}</span>)}</div>
            </div>
          )}

          {/* 박스 선택 */}
          {selRack&&boxes.length>0&&(
            <div style={{marginBottom:16}}>
              <div style={{fontSize:12,color:"#888",fontWeight:600,marginBottom:8}}>📦 박스</div>
              <div>{boxes.map(b=>{
                const cnt=Object.keys(ln2Data[selSheet][selRack][b]?.cells||{}).length;
                return <span key={b} style={S.chip(selBox===b,"#059669")} onClick={()=>setSelBox(b)}>
                  {b} <span style={{fontSize:10,opacity:0.7,marginLeft:3}}>{cnt}칸</span>
                </span>;
              })}</div>
            </div>
          )}

          {/* 그리드 */}
          {cells&&(
            <div>
              <div style={{fontSize:14,fontWeight:700,marginBottom:8}}>
                {selSheet} — 랙 #{selRack} — 박스 {selBox}
              </div>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
                {LEGEND.map(l=><div key={l.label} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#555"}}><div style={{width:10,height:10,borderRadius:3,background:l.color}}/>{l.label}</div>)}
              </div>
              <CellGrid cells={cells} onCellClick={(key,content)=>setEditModal({key,content})}/>
            </div>
          )}
        </div>
      )}

      {subTab==="search"&&(
        <div>
          <input style={S.searchInput} placeholder="🔍 세포주 이름 검색..." value={search} onChange={e=>setSearch(e.target.value)} autoFocus/>
          {search.trim().length>1&&searchResults.length===0&&<div style={S.empty}>검색 결과 없음</div>}
          {searchResults.map((r,i)=>{
            const color=getCellColor(r.content);
            return <div key={i} style={S.resultItem} onClick={()=>{setSelSheet(r.sheet);setSelRack(r.rack);setSelBox(r.box);setSubTab("map");}}>
              <div style={{...S.badge,background:color||"#6b7280"}}>{r.sheet} — 랙#{r.rack} — 박스{r.box} — {r.key}</div>
              <div style={{fontSize:13,color:"#333",marginTop:2}}>{r.content}</div>
            </div>;
          })}
          {search.trim().length<=1&&<div style={S.empty}>2글자 이상 입력하세요</div>}
        </div>
      )}

      {editModal&&(
        <EditModal cellKey={editModal.key} content={editModal.content}
          onSave={(key,val)=>{updateCell(key,val);setEditModal(null);notify("저장됨");}}
          onDelete={(key)=>{updateCell(key,"");setEditModal(null);notify("삭제됨");}}
          onClose={()=>setEditModal(null)}/>
      )}

      {importModal&&(
        <div style={S.modalBg}>
          <div style={S.modal}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:10}}>📂 LN2 엑셀 불러오기</div>
            <div style={{fontSize:13,color:"#555",marginBottom:6}}>
              {Object.entries(importModal.data).map(([sheet,racks])=>(
                <div key={sheet} style={{marginBottom:4}}>
                  <b>{sheet}</b>: 랙 {Object.keys(racks).length}개, 박스 {Object.values(racks).flatMap(r=>Object.keys(r)).length}개
                </div>
              ))}
            </div>
            <div style={{fontSize:13,color:"#666",marginBottom:16,marginTop:10}}>기존 데이터를 어떻게 할까요?</div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end",flexWrap:"wrap"}}>
              <button style={S.btnG} onClick={()=>setImportModal(null)}>취소</button>
              <button style={S.btnO} onClick={()=>{setLn2Data(p=>({...p,...importModal.data}));setImportModal(null);notify("LN2 데이터 추가됨");}}>기존 유지 + 추가</button>
              <button style={S.btnP} onClick={()=>{setLn2Data(importModal.data);setImportModal(null);notify("LN2 데이터 불러옴");}}>전체 교체</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── 주문 관리 ─────────────────────────────────────────────
const ORDER_SAMPLE = [
  {id:1,name:"RPMI 1640 배지",cat:"R8758",maker:"Sigma-Aldrich",link:"https://www.sigmaaldrich.com"},
  {id:2,name:"FBS (Fetal Bovine Serum)",cat:"16000044",maker:"Gibco",link:""},
  {id:3,name:"Trypan Blue Solution",cat:"T8154",maker:"Sigma-Aldrich",link:""},
];

function OrderApp({ notify }) {
  const [db,setDb]=useState(ORDER_SAMPLE);
  const [history,setHistory]=useState([]);
  const [selected,setSelected]=useState({});
  const [search,setSearch]=useState("");
  const [subTab,setSubTab]=useState("order");
  const [form,setForm]=useState({name:"",cat:"",maker:"",link:""});
  const [gmailOpen,setGmailOpen]=useState(false);
  const [gmailTo,setGmailTo]=useState("");
  const [gmailSubject,setGmailSubject]=useState("");
  const [nextId,setNextId]=useState(100);
  const [mapModal,setMapModal]=useState(null);

  const addItem=()=>{ if(!form.name.trim()||!form.cat.trim()){notify("품목명과 카탈로그번호는 필수");return;} setDb(p=>[...p,{id:nextId,...form}]);setNextId(n=>n+1);setForm({name:"",cat:"",maker:"",link:""});notify("추가됨"); };
  const deleteItem=(id)=>{ setDb(p=>p.filter(x=>x.id!==id));setSelected(p=>{const n={...p};delete n[id];return n;});notify("삭제됨"); };
  const toggleSelect=(id)=>setSelected(p=>{const n={...p};if(n[id])delete n[id];else n[id]={qty:1};return n;});
  const updateQty=(id,val)=>setSelected(p=>({...p,[id]:{qty:Math.max(1,parseInt(val)||1)}}));
  const getItems=()=>Object.keys(selected).map(id=>{const item=db.find(x=>x.id==id);return item?{...item,qty:selected[id].qty}:null;}).filter(Boolean);

  const downloadCSV=()=>{
    const items=getItems(); if(!items.length){notify("품목 선택 필요");return;}
    const rows=[["카탈로그번호","품목명","제조사","수량","링크"],...items.map(x=>[x.cat,x.name,x.maker,x.qty,x.link])];
    const csv=rows.map(r=>r.map(c=>`"${String(c||"").replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob=new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
    const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;
    const today=new Date().toISOString().slice(0,10);
    a.download=`주문서_${today}.csv`;a.click();URL.revokeObjectURL(url);
    setHistory(p=>[{date:today,items:items.map(x=>({name:x.name,cat:x.cat,maker:x.maker,link:x.link,qty:x.qty}))}, ...p].slice(0,30));
    notify("CSV 다운로드!");
  };

  const buildBody=()=>{
    const items=getItems();const today=new Date().toLocaleDateString("ko-KR");
    let body=`안녕하세요,\n\n아래 물품 주문 요청드립니다.\n\n주문일: ${today}\n\n${"─".repeat(28)}\n`;
    items.forEach((x,i)=>{body+=`${i+1}. ${x.name}\n   카탈로그번호: ${x.cat}\n`;if(x.maker)body+=`   제조사: ${x.maker}\n`;body+=`   수량: ${x.qty}\n`;if(x.link)body+=`   링크: ${x.link}\n`;body+="\n";});
    return body+`${"─".repeat(28)}\n\n감사합니다.`;
  };

  const handleFileUpload=(e)=>{
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=(evt)=>{
      try{
        const wb=XLSX.read(evt.target.result,{type:"binary"});const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});if(!rows.length){notify("데이터 없음");return;}
        const headers=Object.keys(rows[0]);
        const af=(cands)=>headers.find(h=>cands.some(c=>h.replace(/\s/g,"").toLowerCase().includes(c)))||"";
        const mapping={name:af(["품목명","name","item","품명","제품명"]),cat:af(["카탈로그","catalog","cat","번호"]),maker:af(["제조사","maker","manufacturer"]),link:af(["링크","link","url"])};
        setMapModal({headers,rows,mapping});e.target.value="";
      }catch{notify("파일 읽기 오류");}
    };
    reader.readAsBinaryString(file);
  };

  const applyMapping=()=>{
    const{rows,mapping}=mapModal;let added=0,skipped=0;
    setDb(prev=>{let updated=[...prev];let id=nextId;
      rows.forEach(row=>{const name=String(row[mapping.name]??"").trim();const cat=String(row[mapping.cat]??"").trim();const maker=String(row[mapping.maker]??"").trim();const link=String(row[mapping.link]??"").trim();
        if(!name&&!cat){skipped++;return;}const dup=updated.find(x=>x.cat===cat&&x.name===name);if(dup){skipped++;return;}updated.push({id:id++,name,cat,maker,link});added++;
      });setNextId(id);return updated;
    });setMapModal(null);notify(`${added}개 추가됨${skipped?`, ${skipped}개 제외`:""}`)
  };

  const filtered=db.filter(x=>!search||x.name.toLowerCase().includes(search.toLowerCase())||x.cat.toLowerCase().includes(search.toLowerCase()));
  const selCount=Object.keys(selected).length;

  const S={
    subTabs:{display:"flex",borderBottom:"1.5px solid #e5e7eb",marginBottom:14,gap:2},
    subTab:{padding:"7px 14px",fontSize:13,background:"none",border:"none",borderBottom:"2px solid transparent",marginBottom:-1.5,cursor:"pointer",color:"#999",fontWeight:500},
    subTabActive:{color:"#1a1a1a",borderBottomColor:"#1a1a1a",fontWeight:700},
    si:{width:"100%",padding:"9px 14px",fontSize:14,border:"1.5px solid #e8e8e8",borderRadius:9,marginBottom:12,outline:"none",background:"#fafafa",boxSizing:"border-box"},
    ir:{display:"flex",alignItems:"center",padding:"11px 14px",border:"1.5px solid #ececec",borderRadius:11,background:"#fff",marginBottom:7},
    irS:{borderColor:"#2563eb",background:"#f0f6ff"},
    sb:{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",background:"#f7f7f8",borderRadius:11,marginTop:14,flexWrap:"wrap",gap:8},
    card:{background:"#fff",border:"1.5px solid #ececec",borderRadius:13,padding:"18px 18px 14px",marginBottom:10},
    cardT:{fontSize:15,fontWeight:700,marginBottom:14,color:"#111"},
    inp:{padding:"8px 12px",fontSize:14,border:"1.5px solid #e0e0e0",borderRadius:8,outline:"none",width:"100%",boxSizing:"border-box",background:"#fafafa"},
    lbl:{display:"block",fontSize:12,color:"#888",marginBottom:4,fontWeight:500},
    tag:{fontSize:11,padding:"3px 9px",borderRadius:99,border:"1px solid #e5e5e5",color:"#555",background:"#f5f5f5"},
    empty:{textAlign:"center",padding:"40px",color:"#bbb",fontSize:14},
    btnP:{padding:"8px 16px",fontSize:13,background:"#1a1a1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600},
    btnO:{padding:"8px 14px",fontSize:13,background:"#fff",color:"#333",border:"1.5px solid #ccc",borderRadius:8,cursor:"pointer"},
    btnG:{padding:"8px 14px",fontSize:13,background:"transparent",color:"#777",border:"1.5px solid #e5e5e5",borderRadius:8,cursor:"pointer"},
    btnD:{padding:"5px 12px",fontSize:12,background:"#fff",color:"#e53e3e",border:"1px solid #fca5a5",borderRadius:7,cursor:"pointer"},
    qi:{width:54,padding:"5px 8px",fontSize:13,border:"1.5px solid #ddd",borderRadius:7,textAlign:"center",outline:"none"},
    mb:{position:"fixed",inset:0,background:"rgba(0,0,0,0.32)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
    mo:{background:"#fff",borderRadius:15,padding:"26px 24px 22px",width:"92%",maxWidth:440,boxShadow:"0 8px 40px rgba(0,0,0,0.16)"},
    pv:{background:"#f6f6f6",borderRadius:9,padding:"12px 14px",fontSize:12,whiteSpace:"pre-wrap",lineHeight:1.75,maxHeight:160,overflowY:"auto",color:"#444"},
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
          <input style={S.si} placeholder="🔍  품목명 또는 카탈로그번호..." value={search} onChange={e=>setSearch(e.target.value)}/>
          {filtered.length===0?<div style={S.empty}>품목이 없습니다</div>:
            filtered.map(item=>(
              <div key={item.id} style={{...S.ir,...(selected[item.id]?S.irS:{})}}>
                <input type="checkbox" checked={!!selected[item.id]} onChange={()=>toggleSelect(item.id)} style={{marginRight:10,width:16,height:16,cursor:"pointer",flexShrink:0,accentColor:"#2563eb"}}/>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:14,fontWeight:600,color:"#1a1a1a"}}>{item.name}</div>
                  <div style={{fontSize:12,color:"#999",marginTop:2}}>{item.cat}{item.maker?` · ${item.maker}`:""}{item.link?<a href={item.link} target="_blank" rel="noreferrer" style={{color:"#2563eb",marginLeft:6}}>링크↗</a>:""}</div>
                </div>
                {selected[item.id]&&<div style={{display:"flex",alignItems:"center",gap:5}}><span style={{fontSize:12,color:"#999"}}>수량</span><input type="number" min="1" value={selected[item.id].qty} onChange={e=>updateQty(item.id,e.target.value)} style={S.qi}/></div>}
              </div>
            ))
          }
          <div style={S.sb}>
            <span style={{fontSize:13,color:"#666"}}>{selCount>0?`✅ ${selCount}개 선택됨`:"품목을 선택하세요"}</span>
            <div style={{display:"flex",gap:8}}>
              <button style={S.btnG} onClick={()=>setSelected({})}>초기화</button>
              <button style={S.btnO} onClick={()=>{if(!getItems().length){notify("품목 선택 필요");return;}setGmailSubject(`[주문 요청] 연구실 물품 ${new Date().toLocaleDateString("ko-KR")}`);setGmailOpen(true);}}>✉️ Gmail</button>
              <button style={S.btnP} onClick={downloadCSV}>⬇️ CSV</button>
            </div>
          </div>
        </div>
      )}

      {subTab==="db"&&(
        <div>
          <div style={S.card}>
            <div style={S.cardT}>새 품목 추가</div>
            <div style={{display:"flex",gap:8,marginBottom:9,flexWrap:"wrap"}}>
              <input style={{...S.inp,flex:"1 1 140px"}} placeholder="품목명 *" value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/>
              <input style={{...S.inp,flex:"1 1 120px"}} placeholder="카탈로그번호 *" value={form.cat} onChange={e=>setForm(f=>({...f,cat:e.target.value}))}/>
            </div>
            <div style={{display:"flex",gap:8,marginBottom:13,flexWrap:"wrap"}}>
              <input style={{...S.inp,flex:"1 1 120px"}} placeholder="제조사" value={form.maker} onChange={e=>setForm(f=>({...f,maker:e.target.value}))}/>
              <input style={{...S.inp,flex:"2 1 180px"}} placeholder="링크 (URL)" value={form.link} onChange={e=>setForm(f=>({...f,link:e.target.value}))}/>
            </div>
            <div style={{textAlign:"right"}}><button style={S.btnP} onClick={addItem}>+ 추가</button></div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10,margin:"6px 0 14px",padding:"11px 14px",background:"#f0f6ff",borderRadius:10,border:"1.5px dashed #93c5fd"}}>
            <span style={{fontSize:13,color:"#2563eb",fontWeight:600}}>📂 엑셀 업로드</span>
            <span style={{fontSize:12,color:"#888",flex:1}}>.xlsx/.xls/.csv · 중복 자동 제외</span>
            <label style={{...S.btnP,padding:"6px 14px",cursor:"pointer"}}>파일 선택<input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} style={{display:"none"}}/></label>
          </div>
          <div style={{fontSize:13,fontWeight:600,color:"#666",margin:"14px 0 8px"}}>등록된 품목 ({db.length})</div>
          {db.length===0?<div style={S.empty}>품목 없음</div>:
            db.map(item=>(
              <div key={item.id} style={S.ir}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:14,fontWeight:600}}>{item.name}</div>
                  <div style={{fontSize:12,color:"#999",marginTop:2}}>{item.cat}{item.maker?` · ${item.maker}`:""}{item.link?<a href={item.link} target="_blank" rel="noreferrer" style={{color:"#2563eb",marginLeft:6}}>링크↗</a>:""}</div>
                </div>
                <button style={S.btnD} onClick={()=>deleteItem(item.id)}>삭제</button>
              </div>
            ))
          }
        </div>
      )}

      {subTab==="history"&&(
        <div>
          {history.length===0?<div style={S.empty}>주문 이력 없음<br/><span style={{fontSize:12}}>CSV 다운로드 시 저장됩니다</span></div>:
            history.map((h,idx)=>(
              <div key={idx} style={{padding:"13px 16px",border:"1.5px solid #ececec",borderRadius:11,marginBottom:7,cursor:"pointer",background:"#fff"}} onClick={()=>{const h2=history[idx];const ns={};h2.items.forEach(hi=>{const f=db.find(x=>x.cat===hi.cat&&x.name===hi.name);if(f)ns[f.id]={qty:hi.qty};});setSelected(ns);setSubTab("order");notify(`${h2.date} 내역 불러옴`);}}>
                <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontSize:13,fontWeight:600}}>📅 {h.date}</span><span style={{fontSize:12,color:"#2563eb"}}>재주문 →</span></div>
                <div style={{marginTop:7,display:"flex",flexWrap:"wrap",gap:4}}>{h.items.map((x,i)=><span key={i} style={S.tag}>{x.name} ×{x.qty}</span>)}</div>
              </div>
            ))
          }
        </div>
      )}

      {gmailOpen&&(
        <div style={S.mb} onClick={()=>setGmailOpen(false)}>
          <div style={S.mo} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>✉️ Gmail 초안</div>
            <div style={{marginBottom:10}}><label style={S.lbl}>받는 사람</label><input style={S.inp} placeholder="supplier@example.com" value={gmailTo} onChange={e=>setGmailTo(e.target.value)}/></div>
            <div style={{marginBottom:10}}><label style={S.lbl}>제목</label><input style={S.inp} value={gmailSubject} onChange={e=>setGmailSubject(e.target.value)}/></div>
            <div style={{marginBottom:16}}><label style={S.lbl}>본문 미리보기</label><div style={S.pv}>{buildBody()}</div></div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button style={S.btnG} onClick={()=>setGmailOpen(false)}>취소</button>
              <button style={S.btnP} onClick={()=>{window.open(`https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(gmailTo)}&su=${encodeURIComponent(gmailSubject)}&body=${encodeURIComponent(buildBody())}`,"_blank");setGmailOpen(false);}}>Gmail에서 열기 ↗</button>
            </div>
          </div>
        </div>
      )}

      {mapModal&&(
        <div style={S.mb} onClick={()=>setMapModal(null)}>
          <div style={{...S.mo,maxWidth:480}} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:14}}>📋 컬럼 매핑</div>
            {[{key:"name",label:"품목명 *"},{key:"cat",label:"카탈로그번호 *"},{key:"maker",label:"제조사"},{key:"link",label:"링크"}].map(({key,label})=>(
              <div key={key} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                <span style={{fontSize:13,fontWeight:600,color:"#444",width:110,flexShrink:0}}>{label}</span>
                <select value={mapModal.mapping[key]} onChange={e=>setMapModal(m=>({...m,mapping:{...m.mapping,[key]:e.target.value}}))} style={{...S.inp,background:"#fff",flex:1}}>
                  <option value="">(매핑 안 함)</option>
                  {mapModal.headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
            <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:8}}>
              <button style={S.btnG} onClick={()=>setMapModal(null)}>취소</button>
              <button style={S.btnP} onClick={applyMapping}>가져오기</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── 메인 ─────────────────────────────────────────────────
export default function App() {
  const [mainTab, setMainTab] = useState("order");
  const [stockTab, setStockTab] = useState("df");
  const [notif, setNotif] = useState("");
  const notify = useCallback((msg)=>{setNotif(msg);setTimeout(()=>setNotif(""),2400);},[]);

  // 전역 상태 — 탭 이동해도 유지
  const [dfBoxes, setDfBoxes] = useState(()=>{
    try { const s=localStorage.getItem("df_boxes"); return s?JSON.parse(s):INIT_STORAGE.deepfreezer; } catch { return INIT_STORAGE.deepfreezer; }
  });
  const [ln2Data, setLn2Data] = useState(()=>{
    try { const s=localStorage.getItem("ln2_data"); return s?JSON.parse(s):{}; } catch { return {}; }
  });

  // localStorage 동기화
  const updateDfBoxes = useCallback((fn) => {
    setDfBoxes(prev => {
      const next = typeof fn === "function" ? fn(prev) : fn;
      try { localStorage.setItem("df_boxes", JSON.stringify(next)); } catch {}
      return next;
    });
  }, []);

  const updateLn2Data = useCallback((fn) => {
    setLn2Data(prev => {
      const next = typeof fn === "function" ? fn(prev) : fn;
      try { localStorage.setItem("ln2_data", JSON.stringify(next)); } catch {}
      return next;
    });
  }, []);

  const S = {
    app:{fontFamily:"'Apple SD Gothic Neo','Malgun Gothic',sans-serif",maxWidth:860,margin:"0 auto",padding:"16px 16px 60px"},
    header:{padding:"14px 0 14px",display:"flex",alignItems:"center",borderBottom:"1.5px solid #f0f0f0",marginBottom:0},
    mainTabs:{display:"flex",background:"#f5f5f5",borderRadius:12,padding:4,gap:4,margin:"16px 0 20px"},
    mainTab:{flex:1,padding:"10px",fontSize:14,background:"transparent",border:"none",borderRadius:9,cursor:"pointer",color:"#888",fontWeight:600},
    mainTabActive:{background:"#fff",color:"#111",boxShadow:"0 1px 4px rgba(0,0,0,0.10)"},
    stockTabs:{display:"flex",background:"#f0f6ff",borderRadius:10,padding:3,gap:3,marginBottom:18},
    stockTab:{flex:1,padding:"8px",fontSize:13,background:"transparent",border:"none",borderRadius:8,cursor:"pointer",color:"#6b99d6",fontWeight:600},
    stockTabActive:{background:"#fff",color:"#1a4b8c",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"},
  };

  return (
    <div style={S.app}>
      <div style={S.header}>
        <span style={{fontSize:19,fontWeight:700,color:"#111"}}>🔬 연구실 통합 관리</span>
      </div>

      <div style={S.mainTabs}>
        <button style={{...S.mainTab,...(mainTab==="order"?S.mainTabActive:{})}} onClick={()=>setMainTab("order")}>🧪 주문 관리</button>
        <button style={{...S.mainTab,...(mainTab==="stock"?S.mainTabActive:{})}} onClick={()=>setMainTab("stock")}>🧊 Cell Stock Map</button>
      </div>

      {mainTab==="order"&&<OrderApp notify={notify}/>}

      {mainTab==="stock"&&(
        <div>
          <div style={S.stockTabs}>
            <button style={{...S.stockTab,...(stockTab==="df"?S.stockTabActive:{})}} onClick={()=>setStockTab("df")}>❄️ Deepfreezer</button>
            <button style={{...S.stockTab,...(stockTab==="ln2"?S.stockTabActive:{})}} onClick={()=>setStockTab("ln2")}>🧊 LN2</button>
          </div>
          {stockTab==="df"&&<DeepfreezerTab notify={notify} boxes={dfBoxes} setBoxes={updateDfBoxes}/>}
          {stockTab==="ln2"&&<LN2Tab notify={notify} ln2Data={ln2Data} setLn2Data={updateLn2Data}/>}
        </div>
      )}

      <Notification msg={notif}/>
    </div>
  );
}