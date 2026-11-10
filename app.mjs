export const headers=['N O','회사명','부서명','사번','성명','READ(R)','READ(L)','SPH(R)','SPH(L)','CYL(R)','CYL(L)','AXIS(R)','AXIS(L)','ADD','OH','PD','FRAME','차광클립','렌즈/옵션','비고','날짜','월','휴대폰','생년'];
export const keys=['no','company','dept','empId','name','readR','readL','sphR','sphL','cylR','cylL','axisR','axisL','add','oh','pd','frame','shadeClip','lensOption','remark','date','month','phone','birthYear'];
const norm=v=>String(v??'').normalize('NFC').replace(/[\s\u200B-\u200D\uFEFF]/g,'');
const eyeOf=s=>{const c=norm(s).toUpperCase().replace(/[().\-_/·,]/g,'');if(/^(OD|RIGHT|RE|R$|R[우안측]|우|오른)/.test(c))return 'od';if(/^(OS|LEFT|LE|L$|L[좌안측]|좌|왼)/.test(c))return 'os';return null;};
const key=x=>[x.center,x.team,x.name].map(norm).join('|');
export const text=c=>{const v=c?.value;if(v==null)return '';if(typeof v==='object'){if(v.richText)return v.richText.map(x=>x.text).join('').trim();if('result'in v)return String(v.result??'').trim();if('text'in v)return String(v.text).trim();}return String(v).trim()};
const num=v=>/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(v)?Number(v):null;
function columns(sheet,row,names){const map={};sheet.getRow(row).eachCell((c,i)=>{const h=text(c).split('\n')[0].trim();if(!(h in map))map[h]=i;});const missing=names.filter(n=>!map[n]);if(missing.length)throw Error(`${sheet.name} ${row}행에 필요한 열이 없습니다: ${missing.join(', ')}`);return map;}
export function parse(wb){const a=wb.worksheets.find(s=>s.name.trim()==='보안경명단'),v=wb.worksheets.find(s=>s.name.trim()==='시력데이터');if(!a||!v)throw Error('한 파일에 보안경명단과 시력데이터 시트가 모두 필요합니다.');
const ac=columns(a,4,['센터','팀','사번','성명','구분']),vc=columns(v,3,['센터명','성명','팀','구분','SPH','CYL','AXIS','P.D','ADD']);const applicants=[],visions=[],issues=[];
for(let i=5;i<=a.rowCount;i++){const r=a.getRow(i),get=n=>text(r.getCell(ac[n]));if(!get('성명')&&!get('사번'))continue;const x={center:get('센터'),team:get('팀'),name:get('성명'),empId:get('사번'),category:get('구분'),row:i};if(!x.name)issues.push({...x,reason:'신청명단 성명 누락',source:'신청명단'});else applicants.push(x);}
let current=null;
const finish=()=>{if(current)visions.push(current);current=null};
for(let i=4;i<=v.rowCount;i++){const r=v.getRow(i),get=n=>vc[n]?text(r.getCell(vc[n])):'',eye=eyeOf(get('구분'));const name=get('성명'),center=get('센터명'),team=get('팀');const values={sph:get('SPH'),cyl:get('CYL'),axis:get('AXIS'),pd:get('P.D'),add:get('ADD')};
if(!name&&!eye&&!Object.values(values).some(Boolean)){finish();continue;}
if(!eye){finish();issues.push({name,center,team,row:i,source:'시력데이터',reason:'우·좌안 구분을 인식할 수 없음'});continue;}
if(current&&((name&&norm(name)!==norm(current.name))||(center&&norm(center)!==norm(current.center))||(team&&norm(team)!==norm(current.team))||current[eye]))finish();
if(!current){current={name,center,team,category:get('도수보안경/돋보기'),row:i};if(!name){issues.push({...current,reason:'독립된 안구 행의 성명 누락',source:'시력데이터'});current=null;continue;}}
current[eye]=values;if(current.od&&current.os)finish();
}finish();return {applicants,visions,issues};}
export function validate(r){const w=[];for(const [k,label]of [['empId','사번'],['dept','센터명'],['name','성명']])if(!r[k].trim())w.push(label+' 누락');for(const side of ['R','L']){if(!r['sph'+side]&&!r['cyl'+side]&&!r['axis'+side])w.push(`${side==='R'?'우안':'좌안'} 도수 누락`);for(const field of ['sph','cyl','axis']){const s=r[field+side];if(s&&num(s)===null)w.push(`${field.toUpperCase()}(${side}) 숫자 확인`);}const ax=num(r['axis'+side]);if(r['axis'+side]&&ax!==null&&(ax<0||ax>180))w.push(`AXIS(${side}) 범위 확인`);}
if(r.pd&&(num(r.pd)===null||num(r.pd)<=0||num(r.pd)>90))w.push('PD 값 확인');if(!r.pd)w.push('PD 누락');if(r.add&&(num(r.add)===null||num(r.add)<0))w.push('ADD 값 확인');return w;}
export function match({applicants,visions,issues=[]}){const output=[],unmatched=[],blocked=[...issues],count=xs=>{const m=new Map;xs.forEach(x=>{const k=key(x);m.set(k,[...(m.get(k)||[]),x])});return m},am=count(applicants),vm=count(visions),used=new Set;
for(const a of applicants){const k=key(a),vs=vm.get(k)||[];if(am.get(k).length>1){blocked.push({...a,source:'신청명단',reason:'신청명단 중복 (센터+팀+성명)'});continue;}if(vs.length>1){blocked.push({...a,source:'시력데이터',reason:'시력데이터 중복 (센터+팀+성명)'});used.add(k);continue;}if(!vs.length){unmatched.push({...a,source:'신청명단',reason:visions.some(v=>norm(v.name)===norm(a.name))?'성명 일치 / 센터·팀 불일치':'시력데이터 없음'});continue;}
used.add(k);const v=vs[0],r=Object.fromEntries(keys.map(k=>[k,'']));Object.assign(r,{no:String(output.length+1),dept:v.center||a.center,empId:a.empId,name:a.name,remark:a.category||v.category,notices:[],ack:false,force:false});for(const [eye,suffix]of [['od','R'],['os','L']])for(const field of ['sph','cyl','axis'])r[field+suffix]=v[eye]?.[field]||'';
const od=v.od||{},os=v.os||{};r.add=od.add||os.add||'';if(od.add&&os.add&&!(num(od.add)!==null&&num(od.add)===num(os.add)))r.notices.push('左右 ADD 불일치: 우안 값 사용');
const p=num(od.pd||''),q=num(os.pd||'');if((p!==null&&p>50)||(q!==null&&q>50)){r.pd=p!==null&&p>50?od.pd:os.pd}else if(p!==null&&q!==null){if(p+q>90){r.pd=od.pd;r.notices.push('PD 합계 90 초과: 우안 값 임시 적용')}else{r.pd=String(Math.round((p+q)*100)/100);r.pdSummed=true}}else{r.pd=od.pd||os.pd||'';r.notices.push('PD 한쪽 누락 또는 비숫자: 양안 PD 확인')}
if(a.category&&v.category&&norm(a.category)!==norm(v.category))r.notices.push('구분 불일치: 신청명단 값 사용');output.push(r);}
for(const v of visions)if(!used.has(key(v))&&!am.has(key(v)))unmatched.push({...v,source:'시력데이터',reason:'신청명단 없음'});
return {output,unmatched,blocked,applicants:applicants.length,visions:visions.length};}
export const warnings=r=>[...validate(r),...(r.ack?[]:r.notices)];
