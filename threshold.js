/*───────────────────────────────────────────────────────────────────────────
  MAIN
───────────────────────────────────────────────────────────────────────────*/
const WEEK_LABEL = 'Week 10'; 
const PREV_WEEK  = 'Week ' + (Number(WEEK_LABEL.match(/\d+/)[0]) - 1);

const COL_PATTERNS = {
  vq : /^VQ\d*/i,
  hw : /^(LAB|HW|A)\d*/i,
  ex : /^(E|EXAM)\d*/i
};
const CUT = { vq: 0.60, hw: 0.70, ex: 0.70 };

/*───────────────────────────────────────────────────────────────────────────
  MAIN
───────────────────────────────────────────────────────────────────────────*/
function Student_Help() {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const gradeS   = ss.getSheetByName('PL Gradebook');
    const logS     = ss.getSheetByName('Weekly Contact Log');
    const staffS   = ss.getSheetByName('Staff');
    const enrollS  = ss.getSheetByName('Enrollment');

    /* staff list ------------------------------------------------------- */
    const staff = staffS
      ? staffS.getRange('A2:A').getValues().flat().filter(String).map(v=>v.trim())
      : ['Akshay','Bella','Zidane','Hui Xin','Roxy','Annapurna'];

    /* roster set ------------------------------------------------------- */
    const normalize = s => String(s).toLowerCase()
                         .replace(/[^a-z\s]/g,'')
                         .trim().split(/\s+/).sort().join(' ');
    const enrolled = new Set(
      enrollS.getRange('A2:A').getValues()
             .flat().filter(String).map(normalize)
    );

    /* score‑column indices -------------------------------------------- */
    const hdr=gradeS.getRange(1,1,1,gradeS.getLastColumn()).getValues()[0];
    const vqIdx=[], hwIdx=[], exIdx=[];
    hdr.forEach((h,i)=>{ const c=i+1;
      if(COL_PATTERNS.vq.test(h)) vqIdx.push(c);
      else if(COL_PATTERNS.hw.test(h)) hwIdx.push(c);
      else if(COL_PATTERNS.ex.test(h)) exIdx.push(c);
    });

    /* current / previous week columns --------------------------------- */
    const colThis = findWeekCol_(logS,WEEK_LABEL);
    let   colPrev = null;
    try { colPrev = findWeekCol_(logS,PREV_WEEK); } catch(_){}

    /* map email → staff from previous week ---------------------------- */
    const lastWeek = new Map();
    if(colPrev){
      const labels=logS.getRange(1,1,logS.getLastRow()).getValues().flat();
      let staffHere=null;
      labels.forEach((v,r)=>{
        if(/[A-Za-z]+'s Students/.test(v)){
          staffHere=v.replace(/'s Students/i,'').trim();
          return;
        }
        const email=String(logS.getRange(r+1,colPrev).getValue()).trim().toLowerCase();
        if(staffHere && email) lastWeek.set(email, staffHere);
      });
    }

    /* set of e‑mails already in this week’s column -------------------- */
    const curEmails = new Set(
      logS.getLastRow()>1
        ? logS.getRange(2,colThis,logS.getLastRow()-1,1)
            .getValues().flat().filter(String).map(e=>e.trim().toLowerCase())
        : []
    );

    /* round‑robin pointer --------------------------------------------- */
    const props = PropertiesService.getDocumentProperties();
    let ptr     = Number(props.getProperty('STAFF_PTR')||0);

    /* counters --------------------------------------------------------- */
    let read=0, skipStaff=0, skipRoster=0, struggling=0, written=0;

    const EMAIL_COL=1, NAME_COL=3;

    /* -------------------- MAIN LOOP ---------------------------------- */
    gradeS.getDataRange().getValues().slice(1).forEach(row=>{
      read++;

      const rawName=row[NAME_COL-1];
      if(!rawName) return;
      if(!enrolled.has(normalize(rawName))){ skipRoster++; return; }

      const roleColIdx = 4; // adjust if "Role" column moves
      const role=String(row[roleColIdx-1]||'').toLowerCase();
      if(role.includes('staff')){ skipStaff++; return; }

      // Average of Video Quizzes & HW
      const vqAvg=avg_(row,vqIdx),
            hwAvg=avg_(row,hwIdx),            
            exMin=exIdx.length?min_(row,exIdx):1;

      if(!( (vqIdx.length&&vqAvg<CUT.vq) ||
            (hwIdx.length&&hwAvg<CUT.hw) ||
            (exIdx.length&&exMin<CUT.ex) )) return;
      struggling++;

      const email=String(row[EMAIL_COL-1]).trim().toLowerCase();
      if(!email) return;                       
      if(curEmails.has(email)) return;       

      const staffName = lastWeek.get(email) || staff[ptr % staff.length];
      if(!lastWeek.has(email)) ptr++;  

      const dest = nextRowForStaff_(logS,staffName,colThis);
      logS.getRange(dest,colThis).setValue(email);
      curEmails.add(email);
      written++;
    });

    props.setProperty('STAFF_PTR',String(ptr % staff.length)); // keep bounded

    Logger.log(`read ${read}, staff ${skipStaff}, off‑roster ${skipRoster}, `
             +`struggling ${struggling}, written ${written}`);

  } catch(err){
    Logger.log('❌ '+err);
    throw err;
  }
}

/* Helpers */
function percentToNum_(x){
  if(typeof x==='number') return x>1?x/100:x;
  const s=String(x).replace('%','').trim();
  return s===''?NaN:Number(s)/100;
}
function avg_(row,idx){
  const v=idx.map(i=>percentToNum_(row[i-1])).filter(n=>!isNaN(n));
  return v.length?v.reduce((a,b)=>a+b)/v.length:1;
}
function min_(row,idx){
  const v=idx.map(i=>percentToNum_(row[i-1])).filter(n=>!isNaN(n));
  return v.length?Math.min(...v):1;
}
function findWeekCol_(sh,label){
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=hdr.findIndex(h=>String(h).trim()===label);
  if(i<0) throw Error('Week header not found: '+label);
  return i+1;
}
function nextRowForStaff_(sh,name,col){
  const norm=s=>String(s).trim().replace(/[‘’]/g,"'").toLowerCase();
  const want=norm(name+"'s Students");
  const labels=sh.getRange(1,1,sh.getLastRow()).getValues().flat();
  const top=labels.findIndex(v=>norm(v)===want)+1;
  if(!top) throw Error('Block not found: '+name);
  let r=top+1;
  while(r<=sh.getLastRow()){
    if(/[A-Za-z]+'s Students/i.test(labels[r-1])) break;
    if(!sh.getRange(r,col).getValue()) return r;
    r++;
  }
  sh.insertRowAfter(r-1);
  return r;
}
