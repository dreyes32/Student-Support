/* BEFORE YOU RUN! -> Increment WEEK_LABEL */
const WEEK_LABEL = 'Week 6'; 

/* Store Each Staff Members Assigned Students */
const PREV_WEEK  = 'Week ' + (Number(WEEK_LABEL.match(/\d+/)[0]) - 1);

/* Create Threshold */
const COL_PATTERNS = {
  vq : /^VQ\d*/i,
  hw : /^(LAB|HW|A)\d*/i,
  ex : /^(EX|EXAM)\d*/i
};
const CUT = { vq: 0.60, hw: 0.70, ex: 0.70 };

/* Main */
function Student_Help() {
  try {
    const sheets = SpreadsheetApp.getActiveSpreadsheet();
    const grades = sheets.getSheetByName('PL Gradebook');
    const w_logs = sheets.getSheetByName('Weekly Contact Log');
    const staff = sheets.getSheetByName('Staff');
    const curr_roster = sheets.getSheetByName('Enrollment');

    const team = staff
      ? staff.getRange('A2:A').getValues().flat().filter(String)
      : ['Akshay','Bella','Zidane','Hui Xin','Roxy','Annapurna'];

    /* Filter Out Students Who've Dropped The Course */
    const normalize = s => String(s).toLowerCase()
        .replace(/[^a-z\s]/g,'')
        .trim().split(/\s+/).sort().join(' ');

    const enrolled = new Set(
        curr_roster.getRange('A2:A').getValues()
             .flat().filter(String).map(normalize)
    );

    /* Get Header Indexes, Create Array of Relative Data */
    const header = grades.getRange(1,1,1,grades.getLastColumn()).getValues()[0];
    const vqIdx=[], hwIdx=[], exIdx=[];

    for (let i = 0; i < header.length; i++) {
        const headerLabel = header[i];
        const col_num = i + 1;
      
        if (COL_PATTERNS.vq.test(headerLabel)) {
          vqIdx.push(col_num);
        } 
        else if (COL_PATTERNS.hw.test(headerLabel)) {
          hwIdx.push(col_num);
        } 
        else if (COL_PATTERNS.ex.test(headerLabel)) {
          exIdx.push(col_num);
        }
      }      
    
    /* Retrieves Current & Previous Struggling Student Column */
    const colThis = findWeekCol_(w_logs,WEEK_LABEL);
    let colPrev = null;
    try { colPrev = findWeekCol_(w_logs,PREV_WEEK); } catch(_){}

    /* Maps Staff : Previously Assigned Student */
    const lastWeek = new Map();
    if(colPrev){
      const labels= w_logs.getRange(1,1,w_logs.getLastRow()).getValues().flat();
      let team_map = null;
      labels.forEach((v,r)=>{
        if(/[A-Za-z]+'s Students/.test(v)){
            team_map = v.replace(/'s Students/i,'').trim();
          return;
        }
        const email = String(w_logs.getRange(r+1,colPrev).getValue()).trim().toLowerCase();
        if(team_map && email) lastWeek.set(email, team_map);
      });
    }

    /* Filter Duplicates, Creates New Set / Quick Search */
    const curEmails = new Set(
        w_logs.getRange(2,colThis,w_logs.getLastRow()-1,1)
          .getValues().flat().filter(String).map(e=>e.trim().toLowerCase())
    );

    /* Create (Staff : Student) Assigner */
    const props = PropertiesService.getDocumentProperties();
    let ptr = Number(props.getProperty('STAFF_PTR')||0);

    /* Counters For Processing */
    let read=0, skipStaff=0, skipRoster=0, struggling=0, written=0;

    const EMAIL_COL=1, NAME_COL=3;

    /* Begin Finding Struggling Students */
    grades.getDataRange().getValues().slice(1).forEach(row=>{
      read++;

        const rawName = row[NAME_COL - 1];
            if (rawName === undefined || rawName === null || rawName === '') {
                return;
            }

        const normalized = normalize(rawName);
            if (!enrolled.has(normalized)) {
                skipRoster++;
                return;
            }

        let role = row[3];
        if (role === undefined || role === null) {
            role = '';
        }
        role = String(role).toLowerCase();
        
        if (role.includes('staff')) {
            skipStaff++;
            return;
        }
        
        const vqAvg = avg_(row, vqIdx);
        
        let hwMin = null;
        if (hwIdx.length > 0) {
            hwMin = min_(row, hwIdx);
        }
        
        let exMin = null;
        if (exIdx.length > 0) {
            exMin = min_(row, exIdx);
        } 
        else {
            exMin = 1;
        }
        
        const isStruggling =
            (vqIdx.length > 0 && vqAvg < CUT.vq) ||
            (hwIdx.length > 0 && hwMin < CUT.hw) ||
            (exIdx.length > 0 && exMin < CUT.ex);
        
        if (!isStruggling) {
            return;
        }
        
        struggling++;

        let email = row[EMAIL_COL - 1];
        email = String(email).trim().toLowerCase();
        
        if (curEmails.has(email)) {
            return;
        }
        
        let staffName = null;
        if (lastWeek.has(email)) {
            staffName = lastWeek.get(email);
        } 
        else {
            const index = ptr % team.length;
            staffName = team[index];
            ptr++;
        }
        
        const dest = nextRowForStaff_(w_logs, staffName, colThis);
        w_logs.getRange(dest, colThis).setValue(email);
        curEmails.add(email);
        written++;  
    });

    props.setProperty('STAFF_PTR',String(ptr));

    Logger.log(`read ${read}, staff ${skipStaff}, off‑roster ${skipRoster}, `
             +`struggling ${struggling}, written ${written}`);

  } catch(err){
    Logger.log('❌ '+err);
    throw err;
  }
}

/* Helper Functions */
function percentToNum_(x){
  if(x===0||x==='0'||x==='0%') return NaN;
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
  const header=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=header.findIndex(h=>String(h).trim()===label);
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
