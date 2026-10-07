const VERSION='1.0.5',STORE='nzPilotLogMobile.v1',XLSX_NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DEFAULT={settings:{pilotName:'',secondName:'',clientNo:'',defaultPic:'Self',other16Label:'AS FLIGHT INSTRUCTOR',other17Label:'IFR CROSS COUNTRY',aircraft:[{type:'GA8',reg:'',engineClass:'single'},{type:'C206',reg:'',engineClass:'single'}],commonDetails:[],lastExcelName:'',lastSyncAt:0},flights:[]};
const $=id=>document.getElementById(id),clone=o=>JSON.parse(JSON.stringify(o));
function load(){try{const x=JSON.parse(localStorage.getItem(STORE)||'null');if(x&&x.settings&&Array.isArray(x.flights)){const d=clone(DEFAULT);d.settings=Object.assign(d.settings,x.settings);d.flights=x.flights;return d}}catch{}return clone(DEFAULT)}let db=load();
function save(){localStorage.setItem(STORE,JSON.stringify(db));renderAll()}
function uid(){return crypto.randomUUID?crypto.randomUUID():Date.now().toString(36)+Math.random().toString(36).slice(2)}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function num(v){const n=Number(v);return Number.isFinite(n)?n:0}function h(v){return (Math.round(num(v)*10)/10).toFixed(1)}
function today(){const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}
function fmtDate(iso){if(!iso)return'';const [y,m,d]=iso.split('-');return `${d}/${m}/${y}`}
function normaliseReg(v){v=String(v||'').trim().toUpperCase();if(/^[A-Z]{3}$/.test(v))return 'ZK-'+v;return v}
function parseTime(v){if(v===null||v===undefined||v==='')return 0;if(typeof v==='number')return Number.isFinite(v)&&v>=0?v:NaN;v=String(v).trim();if(v.includes(':')){const [hh,mm]=v.split(':').map(Number);if(!Number.isFinite(hh)||!Number.isFinite(mm)||mm<0||mm>=60)return NaN;return hh+mm/60}const n=Number(v);return Number.isFinite(n)&&n>=0?n:NaN}
