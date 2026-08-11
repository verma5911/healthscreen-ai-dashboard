import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';

import {

LayoutDashboard, Activity, History as HistoryIcon, BarChart3, MapPin, Bell, FileText, Settings as SettingsIcon,

Info, Menu, X, Thermometer, Wind, Mic, User, Users, Camera, Wifi, WifiOff, CheckCircle2, AlertTriangle,

AlertOctagon, Search, Download, Clock, Radio, Cpu, Gauge as GaugeIcon, ChevronDown, RefreshCw, ShieldCheck,

Filter, ArrowUpDown, ChevronLeft, ChevronRight, Server, Waves, Mail, Phone, Save, Building2, Signal, Sparkles

} from 'lucide-react';

import {

LineChart, Line, AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,

PieChart, Pie, Cell, BarChart, Bar

} from 'recharts';

import * as XLSX from 'xlsx';



/* ============================== DESIGN TOKENS ============================== */

const C = {

navy: '#0A1B33',

navy2: '#0F2545',

navy3: '#15305C',

navyBorder: '#1F3A66',

blue: '#2563EB',

blueLight: '#3B82F6',

blueBg: '#EAF1FE',

bg: '#F2F5FA',

card: '#FFFFFF',

border: '#E3E9F2',

text: '#0F1B2D',

textSub: '#5B6B84',

textMute: '#8B98AC',

green: '#128A4A',

greenBg: '#E9F8EF',

greenBorder: '#BEEACF',

yellow: '#B7791F',

yellowBg: '#FEF6E7',

yellowBorder: '#F5DFA9',

orange: '#C2570C',

orangeBg: '#FEF0E4',

orangeBorder: '#F6CFA6',

red: '#C2281C',

redBg: '#FDECEB',

redBorder: '#F5C4BE',

};



const BASE_LAT = 30.9009;

const BASE_LNG = 75.8573;



/* ============================== HELPERS ============================== */

const rand = (min, max) => Math.random() * (max - min) + min;

const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];

const uid = () => Math.random().toString(36).slice(2, 10);



function aqiCategory(aqi) {

if (aqi <= 50) return { label: 'Good', color: C.green, bg: C.greenBg, border: C.greenBorder };

if (aqi <= 100) return { label: 'Moderate', color: C.yellow, bg: C.yellowBg, border: C.yellowBorder };

if (aqi <= 150) return { label: 'Unhealthy (Sensitive)', color: C.orange, bg: C.orangeBg, border: C.orangeBorder };

if (aqi <= 200) return { label: 'Poor', color: C.orange, bg: C.orangeBg, border: C.orangeBorder };

return { label: 'Severe', color: C.red, bg: C.redBg, border: C.redBorder };

}



function riskStyle(level) {

if (level === 'LOW') return { color: C.green, bg: C.greenBg, border: C.greenBorder };

if (level === 'MEDIUM') return { color: C.yellow, bg: C.yellowBg, border: C.yellowBorder };

return { color: C.red, bg: C.redBg, border: C.redBorder };

}



function screeningMeta(result) {

switch (result) {

case 'NO_ABNORMALITY_DETECTED':

return { label: 'No Abnormality Detected', color: C.green, bg: C.greenBg, border: C.greenBorder, Icon: CheckCircle2 };

case 'REQUIRES_FURTHER_SCREENING':

return { label: 'Requires Further Screening', color: C.yellow, bg: C.yellowBg, border: C.yellowBorder, Icon: AlertTriangle };

case 'POSSIBLE_ABNORMALITY':

return { label: 'Possible Abnormality', color: C.orange, bg: C.orangeBg, border: C.orangeBorder, Icon: AlertTriangle };

default:

return { label: 'High Risk Screening', color: C.red, bg: C.redBg, border: C.redBorder, Icon: AlertOctagon };

}

}



function fmtTime(d) {

return new Date(d).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });

}

function fmtDate(d) {

return new Date(d).toLocaleDateString('en-IN', { year: 'numeric', month: 'short', day: '2-digit' });

}

function fmtDateTimeShort(d) {

return new Date(d).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', hour12: false });

}



/* ============================== MOCK DATA ENGINE ============================== */

function genReading(offsetMs = 0) {

const roll = Math.random();

const temperature = roll > 0.9 ? rand(38.2, 39.1) : rand(36.2, 37.6);

const aqi = Math.round(roll > 0.82 ? rand(150, 260) : rand(25, 120));

const cough_score = roll > 0.88 ? rand(0.5, 0.9) : rand(0.02, 0.35);

const breathing_status = roll > 0.92 ? pick(['Irregular', 'Requires Further Screening']) : 'Normal';

const person_detected = Math.random() > 0.04;

const face_detected = person_detected && Math.random() > 0.07;

const num_persons = person_detected ? (Math.random() > 0.85 ? 2 : 1) : 0;



let riskScore = 0;

if (temperature > 37.5) riskScore += 1;

if (temperature > 38.2) riskScore += 2;

if (aqi > 150) riskScore += 1;

if (cough_score > 0.5) riskScore += 2;

if (breathing_status !== 'Normal') riskScore += 2;



let screening_result, risk_level, confidence;

if (riskScore >= 5) {

screening_result = 'HIGH_RISK_SCREENING'; risk_level = 'HIGH'; confidence = rand(78, 96);

} else if (riskScore >= 3) {

screening_result = 'POSSIBLE_ABNORMALITY'; risk_level = 'MEDIUM'; confidence = rand(70, 90);

} else if (riskScore >= 1) {

screening_result = 'REQUIRES_FURTHER_SCREENING'; risk_level = 'LOW'; confidence = rand(72, 88);

} else {

screening_result = 'NO_ABNORMALITY_DETECTED'; risk_level = 'LOW'; confidence = rand(84, 98);

}



const timestamp = new Date(Date.now() - offsetMs).toISOString();



return {

id: uid(),

    timestamp,

    person_detected,

    face_detected,

    num_persons,

temperature: +temperature.toFixed(1),

    aqi,

cough_score: +cough_score.toFixed(2),

    breathing_status,

latitude: +(BASE_LAT + rand(-0.0018, 0.0018)).toFixed(6),

longitude: +(BASE_LNG + rand(-0.0018, 0.0018)).toFixed(6),

    screening_result,

confidence: Math.round(confidence),

    risk_level,

esp32cam_status: Math.random() > 0.05 ? 'ONLINE' : 'OFFLINE',

thermal_status: Math.random() > 0.02 ? 'ONLINE' : 'OFFLINE',

microphone_status: Math.random() > 0.04 ? 'ONLINE' : 'OFFLINE',

aqi_status: Math.random() > 0.02 ? 'ONLINE' : 'OFFLINE',

gps_status: Math.random() > 0.02 ? 'ONLINE' : 'OFFLINE',

};

}



function seedHistory(n) {

const rows = [];

for (let i = 0; i < n; i++) {

rows.push(genReading(i * 5 * 60 * 1000 + Math.round(rand(0, 60000))));

}

return rows;

}



function alertFromReading(r) {

if (r.risk_level === 'HIGH') {

return { id: uid(), type: 'red', title: 'High-Risk Screening', desc: 'Multiple abnormal indicators detected. Further screening recommended.', ts: r.timestamp, source: 'AI Fusion Engine' };

}

if (r.risk_level === 'MEDIUM') {

return { id: uid(), type: 'orange', title: 'Abnormal Sensor Pattern', desc: 'Combined sensor readings fall outside expected screening range.', ts: r.timestamp, source: 'Multi-Sensor Fusion' };

}

if (r.temperature > 37.5) {

`return { id: uid(), type: 'yellow', title: 'Slightly Elevated Temperature', desc: `Thermal reading of ${r.temperature.toFixed(1)}\u00b0C requires attention.`, ts: r.timestamp, source: 'Thermal Sensor' };`

}

if (r.aqi > 150) {

`return { id: uid(), type: 'yellow', title: 'Elevated AQI', desc: `Local air quality index reached ${r.aqi}.`, ts: r.timestamp, source: 'AQI Sensor' };`

}

return { id: uid(), type: 'green', title: 'Normal Parameters', desc: 'All parameters are within configured screening range.', ts: r.timestamp, source: 'Multi-Sensor Fusion' };

}



/* ============================== SMALL UI PRIMITIVES ============================== */

function Pill({ color, bg, border, children, icon: Icon }) {

return (

<span

className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium"

`style={{ color, backgroundColor: bg, border: `1px solid ${border}` }}`

>

{Icon && <Icon size={13} />}

{children}

    </span>

);

}



function Card({ children, className = '', style = {} }) {

return (

<div

`className={`rounded-2xl p-4 sm:p-5 ${className}`}`

`style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, boxShadow: '0 1px 2px rgba(15,27,45,0.04)', ...style }}`

>

{children}

    </div>

);

}



function SectionTitle({ children, action }) {

return (

<div className="flex items-center justify-between mb-3">

<h2 className="text-sm sm:text-base font-semibold tracking-wide" style={{ color: C.text }}>{children}</h2>

{action}

    </div>

);

}



function StatCard({ icon: Icon, iconColor, label, value, valueUnit, sub, statusPill }) {

return (

<Card className="flex flex-col justify-between h-full">

<div className="flex items-start justify-between">

<span className="text-xs font-medium" style={{ color: C.textSub }}>{label}</span>

`<div className="w-8 h-8 rounded-lg flex items-center justify-center" style={{ backgroundColor: `${iconColor}15` }}>`

<Icon size={16} color={iconColor} />

        </div>

      </div>

<div className="mt-3 flex items-baseline gap-1">

<span className="text-2xl font-semibold" style={{ color: C.text }}>{value}</span>

{valueUnit && <span className="text-sm" style={{ color: C.textSub }}>{valueUnit}</span>}

      </div>

<div className="mt-2 flex items-center justify-between">

<span className="text-xs" style={{ color: C.textMute }}>{sub}</span>

{statusPill}

      </div>

    </Card>

);

}



function ConnBadge({ label, online }) {

return (

<span className="inline-flex items-center gap-1.5 text-xs font-medium px-2 py-1 rounded-lg" style={{ backgroundColor: online ? C.greenBg : C.redBg, color: online ? C.green : C.red }}>

{online ? <Wifi size={12} /> : <WifiOff size={12} />}

<span className="hidden md:inline">{label}</span>

<span>{online ? 'Online' : 'Offline'}</span>

    </span>

);

}



function MiniGauge({ value, max, color, label, unit, size = 108 }) {

const pct = Math.min(1, Math.max(0, value / max));

const r = 44;

const c = 2 * Math.PI * r;

return (

<div className="flex flex-col items-center">

<svg width={size} height={size} viewBox="0 0 108 108">

<circle cx="54" cy="54" r={r} fill="none" stroke={C.border} strokeWidth="10" />

<circle

cx="54" cy="54" r={r} fill="none" stroke={color} strokeWidth="10"

strokeDasharray={c} strokeDashoffset={c * (1 - pct)} strokeLinecap="round"

transform="rotate(-90 54 54)"

        />

<text x="54" y="50" textAnchor="middle" fontSize="18" fontWeight="600" fill={C.text}>{value}</text>

<text x="54" y="66" textAnchor="middle" fontSize="10" fill={C.textMute}>{unit}</text>

      </svg>

<span className="text-xs font-medium mt-1" style={{ color: C.textSub }}>{label}</span>

    </div>

);

}



function Toggle({ checked, onChange }) {

return (

<button

onClick={() => onChange(!checked)}

className="w-10 h-6 rounded-full relative transition-colors flex-shrink-0"

style={{ backgroundColor: checked ? C.blue : '#CBD5E1' }}

>

<span className="absolute top-0.5 left-0.5 w-5 h-5 rounded-full bg-white transition-transform" style={{ transform: checked ? 'translateX(16px)' : 'translateX(0)' }} />

    </button>

);

}



const NAV = [

{ key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },

{ key: 'live', label: 'Live Monitoring', icon: Activity },

{ key: 'history', label: 'History & Logs', icon: HistoryIcon },

{ key: 'analytics', label: 'Analytics', icon: BarChart3 },

{ key: 'map', label: 'Map / Location', icon: MapPin },

{ key: 'alerts', label: 'Alerts', icon: Bell },

{ key: 'reports', label: 'Reports', icon: FileText },

{ key: 'settings', label: 'Settings', icon: SettingsIcon },

{ key: 'about', label: 'About Project', icon: Info },

];



/* ============================== SIDEBAR ============================== */

function Sidebar({ page, setPage, mobileOpen, setMobileOpen, alertCount }) {

return (

<>

{mobileOpen && (

<div className="fixed inset-0 bg-black/40 z-40 lg:hidden" onClick={() => setMobileOpen(false)} />

)}

<aside

`className={`fixed z-50 top-0 left-0 h-full w-64 flex flex-col transition-transform duration-200 lg:translate-x-0 ${mobileOpen ? 'translate-x-0' : '-translate-x-full'}`}`

style={{ backgroundColor: C.navy }}

>

`<div className="flex items-center gap-2.5 px-5 py-5" style={{ borderBottom: `1px solid ${C.navyBorder}` }}>`

<div className="w-9 h-9 rounded-lg flex items-center justify-center" style={{ backgroundColor: C.blue }}>

<Sparkles size={18} color="#fff" />

          </div>

<div>

<p className="text-white text-sm font-semibold leading-tight">HealthScreen AI</p>

<p className="text-[11px] leading-tight" style={{ color: '#7C90B3' }}>Multi-Sensor Framework</p>

          </div>

<button className="ml-auto lg:hidden" onClick={() => setMobileOpen(false)}>

<X size={18} color="#9FB1CC" />

          </button>

        </div>



<nav className="flex-1 overflow-y-auto py-3 px-3">

{NAV.map((item) => {

const active = page === item.key;

const Icon = item.icon;

return (

<button

key={item.key}

onClick={() => { setPage(item.key); setMobileOpen(false); }}

className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg mb-1 text-sm font-medium transition-colors relative"

style={{

backgroundColor: active ? C.blue : 'transparent',

color: active ? '#fff' : '#A9BAD6',

}}

>

<Icon size={17} />

<span className="flex-1 text-left">{item.label}</span>

{item.key === 'alerts' && alertCount > 0 && (

<span className="text-[10px] font-semibold rounded-full px-1.5 py-0.5" style={{ backgroundColor: active ? 'rgba(255,255,255,0.25)' : C.red, color: '#fff' }}>

{alertCount}

                  </span>

)}

              </button>

);

})}

        </nav>



`<div className="px-4 py-4 text-[11px] leading-snug" style={{ borderTop: `1px solid ${C.navyBorder}`, color: '#6F84A8' }}>`

          Raspberry Pi 3B+ &middot; ESP32-CAM

<br />Research prototype build

        </div>

      </aside>

    </>

);

}



/* ============================== HEADER ============================== */

function Header({ setMobileOpen, live, now, demoMode, setDemoMode, systemOnline, rpiOnline }) {

return (

`<header className="sticky top-0 z-30 flex items-center gap-3 px-4 sm:px-6 py-3" style={{ backgroundColor: C.card, borderBottom: `1px solid ${C.border}` }}>`

<button className="lg:hidden" onClick={() => setMobileOpen(true)}>

<Menu size={20} color={C.text} />

      </button>

<div className="min-w-0">

<p className="text-sm sm:text-base font-semibold truncate" style={{ color: C.text }}>Adaptive Multi-Sensor Health Screening Framework</p>

<p className="text-[11px] sm:text-xs truncate" style={{ color: C.textSub }}>For Sustainable Urban Communities</p>

      </div>



<div className="hidden md:flex items-center gap-2 ml-4">

<ConnBadge label="System" online={systemOnline} />

<ConnBadge label="Raspberry Pi" online={rpiOnline} />

<ConnBadge label="ESP32-CAM" online={live.esp32cam_status === 'ONLINE'} />

      </div>



<div className="ml-auto flex items-center gap-3 sm:gap-4">

<button

onClick={() => setDemoMode(!demoMode)}

className="hidden sm:flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-medium"

`style={{ backgroundColor: demoMode ? C.blueBg : '#F1F3F7', color: demoMode ? C.blue : C.textSub, border: `1px solid ${demoMode ? '#C7DBFB' : C.border}` }}`

>

<RefreshCw size={13} className={demoMode ? 'animate-spin' : ''} style={{ animationDuration: '3s' }} />

          DEMO MODE {demoMode ? 'ON' : 'OFF'}

        </button>

<div className="text-right hidden sm:block">

<p className="text-xs font-medium" style={{ color: C.text }}>{fmtDate(now)}</p>

<p className="text-[11px]" style={{ color: C.textSub }}>{fmtTime(now)}</p>

        </div>

<div className="w-8 h-8 rounded-full flex items-center justify-center" style={{ backgroundColor: C.navy }}>

<User size={15} color="#fff" />

        </div>

      </div>

    </header>

);

}



/* ============================== DASHBOARD PAGE ============================== */

function FusionDiagram({ live }) {

const nodes = [

{ key: 'visual', label: 'Visual', ok: live.face_detected },

{ key: 'thermal', label: 'Thermal', ok: live.thermal_status === 'ONLINE' },

{ key: 'audio', label: 'Audio', ok: live.microphone_status === 'ONLINE' },

{ key: 'aqi', label: 'AQI', ok: live.aqi_status === 'ONLINE' },

{ key: 'gps', label: 'GPS', ok: live.gps_status === 'ONLINE' },

];

const cx = 260, cy = 100, r = 88;

return (

<svg viewBox="0 0 520 200" className="w-full h-auto">

<circle cx={cx} cy={cy} r="30" fill={C.blueBg} stroke={C.blue} strokeWidth="1.5" />

<text x={cx} y={cy - 3} textAnchor="middle" fontSize="10" fontWeight="600" fill={C.blue}>AI FUSION</text>

<text x={cx} y={cy + 10} textAnchor="middle" fontSize="9" fill={C.blue}>ENGINE</text>

{nodes.map((n, i) => {

const angle = (-90 + i * (360 / nodes.length)) * (Math.PI / 180);

const nx = cx + r * Math.cos(angle);

const ny = cy + r * Math.sin(angle);

const color = n.ok ? C.green : C.red;

return (

<g key={n.key}>

<line x1={cx} y1={cy} x2={nx} y2={ny} stroke={n.ok ? '#C7DBFB' : C.redBorder} strokeWidth="1.5" strokeDasharray={n.ok ? '0' : '3 3'} />

<circle cx={nx} cy={ny} r="22" fill="#fff" stroke={color} strokeWidth="1.5" />

<circle cx={nx} cy={ny + 16} r="3.5" fill={color} />

<text x={nx} y={ny + 4} textAnchor="middle" fontSize="9" fontWeight="600" fill={C.text}>{n.label}</text>

          </g>

);

})}

    </svg>

);

}



function TinyChart({ data, dataKey, color, unit, domain }) {

return (

<ResponsiveContainer width="100%" height={140}>

<AreaChart data={data} margin={{ top: 8, right: 8, left: -22, bottom: 0 }}>

<defs>

`<linearGradient id={`grad-${dataKey}`} x1="0" y1="0" x2="0" y2="1">`

<stop offset="0%" stopColor={color} stopOpacity={0.28} />

<stop offset="100%" stopColor={color} stopOpacity={0.02} />

          </linearGradient>

        </defs>

<CartesianGrid stroke={C.border} vertical={false} />

<XAxis dataKey="t" tick={{ fontSize: 9, fill: C.textMute }} tickLine={false} axisLine={{ stroke: C.border }} minTickGap={30} />

<YAxis tick={{ fontSize: 9, fill: C.textMute }} tickLine={false} axisLine={false} width={28} domain={domain || ['auto', 'auto']} />

<Tooltip

`contentStyle={{ fontSize: 11, borderRadius: 8, border: `1px solid ${C.border}` }}`

`formatter={(v) => [`${v}${unit || ''}`, '']}`

`labelFormatter={(l) => `Time ${l}`}`

        />

`<Area type="monotone" dataKey={dataKey} stroke={color} fill={`url(#grad-${dataKey})`} strokeWidth={2} dot={false} isAnimationActive={false} />`

      </AreaChart>

    </ResponsiveContainer>

);

}



function DashboardPage({ live, series, alerts, chartRange, setChartRange }) {

const meta = screeningMeta(live.screening_result);

const aqiMeta = aqiCategory(live.aqi);



const chartData = useMemo(() => {

const rangeCount = { '10m': 20, '1h': 40, '24h': 80, '7d': series.length };

const n = rangeCount[chartRange] || 20;

return series.slice(-n).map((r) => ({

t: fmtDateTimeShort(r.timestamp),

temperature: r.temperature,

aqi: r.aqi,

cough_score: r.cough_score,

confidence: r.confidence,

}));

}, [series, chartRange]);



return (

<div className="space-y-6">

<div>

<SectionTitle>Live Overview</SectionTitle>

<div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">

<StatCard

icon={Users} iconColor={C.blue}

label="Person Detection"

value={live.person_detected ? 'Detected' : 'Not Detected'}

`sub={`${live.num_persons} person(s) \u00b7 Face: ${live.face_detected ? 'Yes' : 'No'}`}`

statusPill={<Pill color={live.person_detected ? C.green : C.textSub} bg={live.person_detected ? C.greenBg : '#F1F3F7'} border={live.person_detected ? C.greenBorder : C.border}>{fmtTime(live.timestamp)}</Pill>}

          />

<StatCard

icon={Thermometer} iconColor={C.red}

label="Temperature"

value={live.temperature.toFixed(1)} valueUnit="\u00b0C"

`sub={`Updated ${fmtTime(live.timestamp)}`}`

statusPill={<Pill {...(live.temperature > 37.5 ? { color: C.orange, bg: C.orangeBg, border: C.orangeBorder } : { color: C.green, bg: C.greenBg, border: C.greenBorder })}>{live.temperature > 37.5 ? 'Elevated' : 'Normal'}</Pill>}

          />

<StatCard

icon={Wind} iconColor={C.blue}

label="AQI"

value={live.aqi}

sub="Environmental status"

statusPill={<Pill color={aqiMeta.color} bg={aqiMeta.bg} border={aqiMeta.border}>{aqiMeta.label}</Pill>}

          />

<StatCard

icon={Mic} iconColor={C.blue}

label="Cough / Audio Score"

value={live.cough_score.toFixed(2)}

`sub={`Updated ${fmtTime(live.timestamp)}`}`

statusPill={<Pill {...(live.cough_score > 0.5 ? { color: C.orange, bg: C.orangeBg, border: C.orangeBorder } : { color: C.green, bg: C.greenBg, border: C.greenBorder })}>{live.cough_score > 0.5 ? 'Elevated' : 'Normal'}</Pill>}

          />

<StatCard

icon={Waves} iconColor={C.blue}

label="Breathing Pattern"

value={live.breathing_status}

sub="Respiratory pattern status"

statusPill={<Pill {...(live.breathing_status === 'Normal' ? { color: C.green, bg: C.greenBg, border: C.greenBorder } : { color: C.orange, bg: C.orangeBg, border: C.orangeBorder })}>{live.breathing_status === 'Normal' ? 'Normal' : 'Attention'}</Pill>}

          />

<StatCard

icon={MapPin} iconColor={C.blue}

label="GPS"

value={live.latitude.toFixed(3)}

`valueUnit={`, ${live.longitude.toFixed(3)}`}`

sub="Screening location"

statusPill={<Pill color={live.gps_status === 'ONLINE' ? C.green : C.red} bg={live.gps_status === 'ONLINE' ? C.greenBg : C.redBg} border={live.gps_status === 'ONLINE' ? C.greenBorder : C.redBorder}>{live.gps_status}</Pill>}

          />

        </div>

      </div>



<div className="grid grid-cols-1 lg:grid-cols-3 gap-4">

<Card className="lg:col-span-2" style={{ borderColor: meta.border }}>

<SectionTitle>AI/ML Screening Result</SectionTitle>

<div className="flex flex-col sm:flex-row sm:items-center gap-4">

<div className="w-16 h-16 rounded-2xl flex items-center justify-center flex-shrink-0" style={{ backgroundColor: meta.bg }}>

<meta.Icon size={30} color={meta.color} />

            </div>

<div className="flex-1">

<p className="text-lg sm:text-xl font-semibold" style={{ color: meta.color }}>{meta.label.toUpperCase()}</p>

<div className="flex flex-wrap items-center gap-x-6 gap-y-1 mt-1.5 text-xs" style={{ color: C.textSub }}>

<span>Confidence: <b style={{ color: C.text }}>{live.confidence}%</b></span>

<span>Risk Level: <b style={{ color: riskStyle(live.risk_level).color }}>{live.risk_level}</b></span>

<span>Timestamp: {fmtTime(live.timestamp)}</span>

<span className="inline-flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: C.green }} />Model Active</span>

              </div>

<div className="w-full h-2 rounded-full mt-3" style={{ backgroundColor: C.border }}>

`<div className="h-2 rounded-full" style={{ width: `${live.confidence}%`, backgroundColor: meta.color }} />`

              </div>

            </div>

          </div>

<p className="text-[11px] mt-3" style={{ color: C.textMute }}>

            Screening indicators only. This system does not diagnose disease.

          </p>

        </Card>



<Card>

<SectionTitle>GPS Location</SectionTitle>

<div className="rounded-xl h-28 relative overflow-hidden mb-3" style={{ backgroundColor: '#EAF1FB' }}>

<svg width="100%" height="100%" viewBox="0 0 200 112">

`{Array.from({ length: 8 }).map((_, i) => <line key={`v${i}`} x1={i * 25} y1="0" x2={i * 25} y2="112" stroke="#D6E3F7" strokeWidth="1" />)}`

`{Array.from({ length: 5 }).map((_, i) => <line key={`h${i}`} x1="0" y1={i * 25} x2="200" y2={i * 25} stroke="#D6E3F7" strokeWidth="1" />)}`

<circle cx="100" cy="56" r="7" fill={C.blue} opacity="0.25" />

<circle cx="100" cy="56" r="4" fill={C.blue} stroke="#fff" strokeWidth="1.5" />

            </svg>

          </div>

<div className="text-xs space-y-1" style={{ color: C.textSub }}>

<div className="flex justify-between"><span>Latitude</span><b style={{ color: C.text }}>{live.latitude.toFixed(6)}</b></div>

<div className="flex justify-between"><span>Longitude</span><b style={{ color: C.text }}>{live.longitude.toFixed(6)}</b></div>

<div className="flex justify-between"><span>Status</span><Pill color={C.green} bg={C.greenBg} border={C.greenBorder}>Locked</Pill></div>

          </div>

        </Card>

      </div>



<Card>

<SectionTitle action={<Pill color={C.green} bg={C.greenBg} border={C.greenBorder}>Data Fusion Status: ACTIVE</Pill>}>Multi-Modal Data Fusion</SectionTitle>

<div className="grid grid-cols-1 lg:grid-cols-2 gap-4 items-center">

<FusionDiagram live={live} />

<div className="grid grid-cols-2 sm:grid-cols-3 gap-3">

{[

{ label: 'Visual / Face', status: live.face_detected ? 'Active' : 'Idle', ok: live.face_detected, value: live.face_detected ? 'Face locked' : 'Scanning' },

`{ label: 'Thermal', status: live.thermal_status, ok: live.thermal_status === 'ONLINE', value: `${live.temperature.toFixed(1)}\u00b0C` },`

{ label: 'Audio / Cough', status: live.microphone_status, ok: live.microphone_status === 'ONLINE', value: live.cough_score.toFixed(2) },

`{ label: 'AQI / Env', status: live.aqi_status, ok: live.aqi_status === 'ONLINE', value: `${live.aqi} AQI` },`

{ label: 'GPS', status: live.gps_status, ok: live.gps_status === 'ONLINE', value: 'Locked' },

].map((s) => (

`<div key={s.label} className="rounded-xl p-3" style={{ border: `1px solid ${C.border}` }}>`

<div className="flex items-center justify-between mb-1">

<span className="text-[11px] font-medium" style={{ color: C.textSub }}>{s.label}</span>

<span className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: s.ok ? C.green : C.red }} />

                </div>

<p className="text-sm font-semibold" style={{ color: C.text }}>{s.value}</p>

<p className="text-[10px]" style={{ color: C.textMute }}>{s.status}</p>

              </div>

))}

          </div>

        </div>

      </Card>



<div>

<SectionTitle action={

<div className="flex gap-1 rounded-lg p-0.5" style={{ backgroundColor: '#F1F3F7' }}>

{['10m', '1h', '24h', '7d'].map((r) => (

<button key={r} onClick={() => setChartRange(r)} className="text-[11px] font-medium px-2.5 py-1 rounded-md" style={{ backgroundColor: chartRange === r ? C.card : 'transparent', color: chartRange === r ? C.blue : C.textSub, boxShadow: chartRange === r ? '0 1px 2px rgba(0,0,0,0.08)' : 'none' }}>

                Last {r}

              </button>

))}

          </div>

}>Real-Time Graphs</SectionTitle>

<div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>Temperature vs Time (\u00b0C)</p><TinyChart data={chartData} dataKey="temperature" color={C.red} unit="\u00b0C" /></Card>

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>AQI vs Time</p><TinyChart data={chartData} dataKey="aqi" color={C.blue} unit="" /></Card>

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>Cough / Audio Score vs Time</p><TinyChart data={chartData} dataKey="cough_score" color={C.orange} unit="" domain={[0, 1]} /></Card>

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>AI Confidence vs Time (%)</p><TinyChart data={chartData} dataKey="confidence" color={C.green} unit="%" domain={[0, 100]} /></Card>

<Card className="sm:col-span-2 xl:col-span-1">

<p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>Breathing Pattern vs Time</p>

<div className="flex flex-wrap gap-1.5 mt-3">

{series.slice(-24).map((r) => (

`<span key={r.id} title={`${fmtDateTimeShort(r.timestamp)} \u2014 ${r.breathing_status}`} className="w-4 h-4 rounded-sm" style={{ backgroundColor: r.breathing_status === 'Normal' ? C.green : C.orange, opacity: 0.85 }} />`

))}

            </div>

<p className="text-[10px] mt-2" style={{ color: C.textMute }}>Each tile = one reading &middot; green = normal, orange = irregular</p>

          </Card>

        </div>

      </div>



<div className="grid grid-cols-1 lg:grid-cols-3 gap-4">

<Card className="lg:col-span-2">

<SectionTitle>Live Camera / ESP32-CAM</SectionTitle>

<div className="rounded-xl aspect-video flex flex-col items-center justify-center" style={{ backgroundColor: C.navy }}>

<Camera size={34} color="#5C79AC" />

<p className="text-xs mt-2" style={{ color: '#8CA0C4' }}>{live.esp32cam_status === 'ONLINE' ? 'Latest captured frame' : 'Stream unavailable \u2014 showing last known frame'}</p>

          </div>

<div className="flex items-center justify-between mt-3 text-xs" style={{ color: C.textSub }}>

<Pill color={live.esp32cam_status === 'ONLINE' ? C.green : C.red} bg={live.esp32cam_status === 'ONLINE' ? C.greenBg : C.redBg} border={live.esp32cam_status === 'ONLINE' ? C.greenBorder : C.redBorder}>{live.esp32cam_status}</Pill>

<span>Last Image Received: {fmtTime(live.timestamp)}</span>

          </div>

        </Card>



<Card>

<SectionTitle>Recent Alerts</SectionTitle>

<div className="space-y-2.5 max-h-64 overflow-y-auto pr-1">

{alerts.slice(0, 6).map((a) => <AlertRow key={a.id} a={a} compact />)}

          </div>

        </Card>

      </div>

    </div>

);

}



/* ============================== LIVE MONITORING PAGE ============================== */

function LiveMonitoringPage({ live }) {

const aqiMeta = aqiCategory(live.aqi);

return (

<div className="space-y-6">

<Card>

<SectionTitle action={<Pill color={C.green} bg={C.greenBg} border={C.greenBorder}>Streaming</Pill>}>Live Sensor Gauges</SectionTitle>

<div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4 justify-items-center">

<MiniGauge value={live.temperature.toFixed(1)} max={40} color={C.red} label="Temperature" unit="\u00b0C" />

<MiniGauge value={live.aqi} max={300} color={aqiMeta.color} label="AQI" unit={aqiMeta.label} />

<MiniGauge value={live.cough_score.toFixed(2)} max={1} color={C.orange} label="Cough Score" unit="index" />

<MiniGauge value={live.confidence} max={100} color={C.blue} label="AI Confidence" unit="%" />

<MiniGauge value={live.num_persons} max={4} color={C.green} label="Persons" unit="count" />

        </div>

      </Card>



<div className="grid grid-cols-1 lg:grid-cols-2 gap-4">

<Card>

<SectionTitle>ESP32-CAM Feed</SectionTitle>

<div className="rounded-xl aspect-video flex flex-col items-center justify-center" style={{ backgroundColor: C.navy }}>

<Camera size={40} color="#5C79AC" />

<p className="text-xs mt-2" style={{ color: '#8CA0C4' }}>Frame @ {fmtTime(live.timestamp)}</p>

<span className="mt-2 text-[10px] font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: live.person_detected ? C.greenBg : '#26456F', color: live.person_detected ? C.green : '#8CA0C4' }}>

{live.person_detected ? 'Person in frame' : 'No person in frame'}

            </span>

          </div>

        </Card>



<Card>

<SectionTitle>Sensor Node Status</SectionTitle>

<div className="space-y-2">

{[

{ label: 'Thermal Sensor', status: live.thermal_status, icon: Thermometer },

{ label: 'Microphone Array', status: live.microphone_status, icon: Mic },

{ label: 'AQI Sensor', status: live.aqi_status, icon: Wind },

{ label: 'GPS Module', status: live.gps_status, icon: MapPin },

{ label: 'ESP32-CAM', status: live.esp32cam_status, icon: Camera },

].map((s) => (

<div key={s.label} className="flex items-center justify-between py-2 px-3 rounded-lg" style={{ backgroundColor: C.bg }}>

<div className="flex items-center gap-2.5">

<s.icon size={15} color={C.textSub} />

<span className="text-sm" style={{ color: C.text }}>{s.label}</span>

                </div>

<Pill color={s.status === 'ONLINE' ? C.green : C.red} bg={s.status === 'ONLINE' ? C.greenBg : C.redBg} border={s.status === 'ONLINE' ? C.greenBorder : C.redBorder}>{s.status}</Pill>

              </div>

))}

          </div>

        </Card>

      </div>



<Card>

<SectionTitle>Breathing Pattern</SectionTitle>

<div className="flex items-center gap-4">

<div className="w-14 h-14 rounded-2xl flex items-center justify-center" style={{ backgroundColor: live.breathing_status === 'Normal' ? C.greenBg : C.orangeBg }}>

<Waves size={26} color={live.breathing_status === 'Normal' ? C.green : C.orange} />

          </div>

<div>

<p className="text-base font-semibold" style={{ color: C.text }}>{live.breathing_status}</p>

<p className="text-xs" style={{ color: C.textSub }}>Respiratory pattern status &middot; updated {fmtTime(live.timestamp)}</p>

          </div>

        </div>

      </Card>

    </div>

);

}



/* ============================== ALERT ROW ============================== */

function alertColor(type) {

if (type === 'green') return { color: C.green, bg: C.greenBg, border: C.greenBorder, Icon: CheckCircle2 };

if (type === 'yellow') return { color: C.yellow, bg: C.yellowBg, border: C.yellowBorder, Icon: AlertTriangle };

if (type === 'orange') return { color: C.orange, bg: C.orangeBg, border: C.orangeBorder, Icon: AlertTriangle };

return { color: C.red, bg: C.redBg, border: C.redBorder, Icon: AlertOctagon };

}

function AlertRow({ a, compact }) {

const m = alertColor(a.type);

return (

`<div className="flex items-start gap-3 rounded-xl p-3" style={{ backgroundColor: m.bg, border: `1px solid ${m.border}` }}>`

<m.Icon size={16} color={m.color} className="mt-0.5 flex-shrink-0" />

<div className="min-w-0 flex-1">

<p className="text-sm font-medium" style={{ color: m.color }}>{a.title}</p>

{!compact && <p className="text-xs mt-0.5" style={{ color: C.textSub }}>{a.desc}</p>}

<div className="flex items-center gap-2 mt-1 text-[10px]" style={{ color: C.textMute }}>

<Clock size={10} /> {fmtDate(a.ts)} {fmtDateTimeShort(a.ts)} <span>&middot;</span> {a.source}

        </div>

      </div>

    </div>

);

}



/* ============================== HISTORY & LOGS PAGE ============================== */

function HistoryPage({ history }) {

const [q, setQ] = useState('');

const [riskFilter, setRiskFilter] = useState('ALL');

const [resultFilter, setResultFilter] = useState('ALL');

const [dateFilter, setDateFilter] = useState('');

const [sortKey, setSortKey] = useState('timestamp');

const [sortDir, setSortDir] = useState('desc');

const [page, setPage] = useState(1);

const pageSize = 10;



const filtered = useMemo(() => {

let rows = history.filter((r) => {

if (riskFilter !== 'ALL' && r.risk_level !== riskFilter) return false;

if (resultFilter !== 'ALL' && r.screening_result !== resultFilter) return false;

if (dateFilter && fmtDate(r.timestamp) !== dateFilter) return false;

if (q) {

`const s = `${r.id} ${r.screening_result} ${r.risk_level} ${r.breathing_status}`.toLowerCase();`

if (!s.includes(q.toLowerCase())) return false;

}

return true;

});

rows.sort((a, b) => {

const av = a[sortKey], bv = b[sortKey];

if (av < bv) return sortDir === 'asc' ? -1 : 1;

if (av > bv) return sortDir === 'asc' ? 1 : -1;

return 0;

});

return rows;

}, [history, q, riskFilter, resultFilter, dateFilter, sortKey, sortDir]);



const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));

const pageRows = filtered.slice((page - 1) * pageSize, page * pageSize);



const toggleSort = (key) => {

if (sortKey === key) setSortDir(sortDir === 'asc' ? 'desc' : 'asc');

else { setSortKey(key); setSortDir('desc'); }

};



const cols = [

{ key: 'timestamp', label: 'Date / Time' },

{ key: 'id', label: 'Session ID' },

{ key: 'temperature', label: 'Temp (\u00b0C)' },

{ key: 'aqi', label: 'AQI' },

{ key: 'cough_score', label: 'Cough Score' },

{ key: 'breathing_status', label: 'Breathing' },

{ key: 'face_detected', label: 'Face' },

{ key: 'screening_result', label: 'Result' },

{ key: 'confidence', label: 'Confidence' },

{ key: 'risk_level', label: 'Risk' },

{ key: 'gps', label: 'GPS' },

];



return (

<div className="space-y-4">

<Card>

<SectionTitle>History & Logs</SectionTitle>

<div className="flex flex-col md:flex-row gap-3 mb-4">

<div className="flex-1 relative">

<Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color={C.textMute} />

<input value={q} onChange={(e) => { setQ(e.target.value); setPage(1); }} placeholder="Search session ID, result, risk..."

`className="w-full pl-9 pr-3 py-2 rounded-lg text-sm outline-none" style={{ border: `1px solid ${C.border}`, color: C.text }} />`

          </div>

`<select value={riskFilter} onChange={(e) => { setRiskFilter(e.target.value); setPage(1); }} className="px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}`, color: C.text }}>`

<option value="ALL">All Risk Levels</option>

<option value="LOW">Low</option>

<option value="MEDIUM">Medium</option>

<option value="HIGH">High</option>

          </select>

`<select value={resultFilter} onChange={(e) => { setResultFilter(e.target.value); setPage(1); }} className="px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}`, color: C.text }}>`

<option value="ALL">All Results</option>

<option value="NO_ABNORMALITY_DETECTED">No Abnormality Detected</option>

<option value="REQUIRES_FURTHER_SCREENING">Requires Further Screening</option>

<option value="POSSIBLE_ABNORMALITY">Possible Abnormality</option>

<option value="HIGH_RISK_SCREENING">High Risk Screening</option>

          </select>

`<input type="date" onChange={(e) => { setDateFilter(e.target.value ? fmtDate(e.target.value) : ''); setPage(1); }} className="px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}`, color: C.text }} />`

        </div>



<div className="overflow-x-auto">

<table className="w-full text-xs" style={{ borderCollapse: 'collapse' }}>

<thead>

<tr style={{ backgroundColor: C.bg }}>

{cols.map((c) => (

<th key={c.key} onClick={() => c.key !== 'gps' && c.key !== 'face_detected' && toggleSort(c.key)} className="text-left px-3 py-2 font-medium whitespace-nowrap cursor-pointer select-none" style={{ color: C.textSub }}>

<span className="inline-flex items-center gap-1">{c.label}{sortKey === c.key && <ArrowUpDown size={10} />}</span>

                  </th>

))}

              </tr>

            </thead>

<tbody>

{pageRows.map((r) => {

const meta = screeningMeta(r.screening_result);

return (

`<tr key={r.id} style={{ borderTop: `1px solid ${C.border}` }}>`

<td className="px-3 py-2 whitespace-nowrap" style={{ color: C.text }}>{fmtDate(r.timestamp)} {fmtDateTimeShort(r.timestamp)}</td>

<td className="px-3 py-2 font-mono" style={{ color: C.textSub }}>{r.id}</td>

<td className="px-3 py-2" style={{ color: C.text }}>{r.temperature.toFixed(1)}</td>

<td className="px-3 py-2" style={{ color: C.text }}>{r.aqi}</td>

<td className="px-3 py-2" style={{ color: C.text }}>{r.cough_score.toFixed(2)}</td>

<td className="px-3 py-2" style={{ color: C.text }}>{r.breathing_status}</td>

<td className="px-3 py-2" style={{ color: C.text }}>{r.face_detected ? 'Yes' : 'No'}</td>

<td className="px-3 py-2"><Pill color={meta.color} bg={meta.bg} border={meta.border}>{meta.label}</Pill></td>

<td className="px-3 py-2" style={{ color: C.text }}>{r.confidence}%</td>

<td className="px-3 py-2"><Pill {...riskStyle(r.risk_level)}>{r.risk_level}</Pill></td>

<td className="px-3 py-2 whitespace-nowrap" style={{ color: C.textSub }}>{r.latitude.toFixed(3)}, {r.longitude.toFixed(3)}</td>

                  </tr>

);

})}

{pageRows.length === 0 && (

<tr><td colSpan={11} className="text-center py-6" style={{ color: C.textMute }}>No matching screening records.</td></tr>

)}

            </tbody>

          </table>

        </div>



<div className="flex items-center justify-between mt-4 text-xs" style={{ color: C.textSub }}>

<span>Showing {pageRows.length ? (page - 1) * pageSize + 1 : 0}\u2013{Math.min(page * pageSize, filtered.length)} of {filtered.length}</span>

<div className="flex items-center gap-1">

`<button disabled={page <= 1} onClick={() => setPage((p) => p - 1)} className="p-1.5 rounded-lg disabled:opacity-30" style={{ border: `1px solid ${C.border}` }}><ChevronLeft size={14} /></button>`

<span className="px-2">{page} / {totalPages}</span>

`<button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)} className="p-1.5 rounded-lg disabled:opacity-30" style={{ border: `1px solid ${C.border}` }}><ChevronRight size={14} /></button>`

          </div>

        </div>

      </Card>

    </div>

);

}



/* ============================== ANALYTICS PAGE ============================== */

const PIE_COLORS = [C.green, C.yellow, C.orange, C.red];

function AnalyticsPage({ history }) {

const stats = useMemo(() => {

const total = history.length;

const normal = history.filter((r) => r.screening_result === 'NO_ABNORMALITY_DETECTED').length;

const further = history.filter((r) => r.screening_result === 'REQUIRES_FURTHER_SCREENING' || r.screening_result === 'POSSIBLE_ABNORMALITY').length;

const high = history.filter((r) => r.screening_result === 'HIGH_RISK_SCREENING').length;

const avg = (key) => total ? (history.reduce((s, r) => s + r[key], 0) / total) : 0;

return {

      total, normal, further, high,

avgTemp: avg('temperature').toFixed(1),

avgAqi: Math.round(avg('aqi')),

avgCough: avg('cough_score').toFixed(2),

avgConf: Math.round(avg('confidence')),

};

}, [history]);



const pieData = [

{ name: 'No Abnormality', value: stats.normal },

{ name: 'Requires Further Screening', value: history.filter((r) => r.screening_result === 'REQUIRES_FURTHER_SCREENING').length },

{ name: 'Possible Abnormality', value: history.filter((r) => r.screening_result === 'POSSIBLE_ABNORMALITY').length },

{ name: 'High Risk', value: stats.high },

];



const trendData = useMemo(() => history.slice().reverse().slice(-30).map((r) => ({

t: fmtDateTimeShort(r.timestamp), temperature: r.temperature, aqi: r.aqi, cough_score: r.cough_score,

})), [history]);



const dailyCounts = useMemo(() => {

const map = {};

history.forEach((r) => { const d = fmtDate(r.timestamp); map[d] = (map[d] || 0) + 1; });

return Object.entries(map).map(([date, count]) => ({ date, count })).slice(-10);

}, [history]);



return (

<div className="space-y-6">

<div className="grid grid-cols-2 lg:grid-cols-4 gap-4">

{[

{ label: 'Total Screenings', value: stats.total, color: C.blue },

{ label: 'Normal Screenings', value: stats.normal, color: C.green },

{ label: 'Further-Screening Cases', value: stats.further, color: C.yellow },

{ label: 'High-Risk Screenings', value: stats.high, color: C.red },

].map((s) => (

<Card key={s.label}>

<p className="text-xs" style={{ color: C.textSub }}>{s.label}</p>

<p className="text-2xl font-semibold mt-1" style={{ color: s.color }}>{s.value}</p>

          </Card>

))}

      </div>

<div className="grid grid-cols-2 lg:grid-cols-4 gap-4">

{[

`{ label: 'Avg Temperature', value: `${stats.avgTemp}\u00b0C` },`

{ label: 'Avg AQI', value: stats.avgAqi },

{ label: 'Avg Cough Score', value: stats.avgCough },

`{ label: 'Avg AI Confidence', value: `${stats.avgConf}%` },`

].map((s) => (

<Card key={s.label}>

<p className="text-xs" style={{ color: C.textSub }}>{s.label}</p>

<p className="text-xl font-semibold mt-1" style={{ color: C.text }}>{s.value}</p>

          </Card>

))}

      </div>



<div className="grid grid-cols-1 lg:grid-cols-2 gap-4">

<Card>

<SectionTitle>Screening Distribution</SectionTitle>

<ResponsiveContainer width="100%" height={230}>

<PieChart>

<Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={50} outerRadius={80} paddingAngle={2}>

{pieData.map((entry, i) => <Cell key={entry.name} fill={PIE_COLORS[i]} />)}

              </Pie>

<Tooltip contentStyle={{ fontSize: 11, borderRadius: 8 }} />

            </PieChart>

          </ResponsiveContainer>

<div className="flex flex-wrap gap-3 justify-center mt-2">

{pieData.map((d, i) => (

<span key={d.name} className="text-[11px] flex items-center gap-1.5" style={{ color: C.textSub }}>

<span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: PIE_COLORS[i] }} />{d.name} ({d.value})

              </span>

))}

          </div>

        </Card>



<Card>

<SectionTitle>Daily Screening Count</SectionTitle>

<ResponsiveContainer width="100%" height={230}>

<BarChart data={dailyCounts} margin={{ left: -20 }}>

<CartesianGrid stroke={C.border} vertical={false} />

<XAxis dataKey="date" tick={{ fontSize: 9, fill: C.textMute }} axisLine={{ stroke: C.border }} tickLine={false} />

<YAxis tick={{ fontSize: 9, fill: C.textMute }} axisLine={false} tickLine={false} />

<Tooltip contentStyle={{ fontSize: 11, borderRadius: 8 }} />

<Bar dataKey="count" fill={C.blue} radius={[4, 4, 0, 0]} />

            </BarChart>

          </ResponsiveContainer>

        </Card>

      </div>



<div className="grid grid-cols-1 sm:grid-cols-3 gap-4">

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>Temperature Trend</p><TinyChart data={trendData} dataKey="temperature" color={C.red} unit="\u00b0C" /></Card>

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>AQI Trend</p><TinyChart data={trendData} dataKey="aqi" color={C.blue} unit="" /></Card>

<Card><p className="text-xs font-medium mb-1" style={{ color: C.textSub }}>Audio / Cough Trend</p><TinyChart data={trendData} dataKey="cough_score" color={C.orange} unit="" domain={[0, 1]} /></Card>

      </div>

    </div>

);

}



/* ============================== MAP PAGE ============================== */

function MapPage({ live, history }) {

const points = history.slice(0, 25);

const proj = (lat, lng) => {

const x = 50 + (lng - BASE_LNG) * 25000;

const y = 50 - (lat - BASE_LAT) * 25000;

return { x: Math.min(96, Math.max(4, x)), y: Math.min(96, Math.max(4, y)) };

};

return (

<div className="space-y-4">

<div className="grid grid-cols-1 lg:grid-cols-3 gap-4">

<Card className="lg:col-span-2">

<SectionTitle>Screening Locations</SectionTitle>

<div className="rounded-xl relative overflow-hidden" style={{ backgroundColor: '#E9F0FB', height: 420 }}>

<svg width="100%" height="100%" viewBox="0 0 100 100" preserveAspectRatio="none">

`{Array.from({ length: 20 }).map((_, i) => <line key={`v${i}`} x1={i * 5} y1="0" x2={i * 5} y2="100" stroke="#D3E1F5" strokeWidth="0.2" />)}`

`{Array.from({ length: 20 }).map((_, i) => <line key={`h${i}`} x1="0" y1={i * 5} x2="100" y2={i * 5} stroke="#D3E1F5" strokeWidth="0.2" />)}`

            </svg>

{points.map((p) => {

const { x, y } = proj(p.latitude, p.longitude);

const rs = riskStyle(p.risk_level);

return (

`<div key={p.id} title={`${fmtDate(p.timestamp)} ${fmtDateTimeShort(p.timestamp)} \u2014 ${p.risk_level}`} className="absolute w-2.5 h-2.5 rounded-full -translate-x-1/2 -translate-y-1/2" style={{ left: `${x}%`, top: `${y}%`, backgroundColor: rs.color, border: '1.5px solid #fff' }} />`

);

})}

{(() => { const { x, y } = proj(live.latitude, live.longitude); return (

`<div className="absolute -translate-x-1/2 -translate-y-1/2 flex flex-col items-center" style={{ left: `${x}%`, top: `${y}%` }}>`

`<div className="w-4 h-4 rounded-full" style={{ backgroundColor: C.blue, border: '2px solid #fff', boxShadow: `0 0 0 6px ${C.blue}22` }} />`

              </div>

); })()}

          </div>

<p className="text-[11px] mt-2" style={{ color: C.textMute }}>Simplified location plot from GPS coordinates. Live deployment renders this on an interactive Leaflet / OpenStreetMap tile layer via the Raspberry Pi API.</p>

        </Card>

<Card>

<SectionTitle>Current GPS Position</SectionTitle>

<div className="space-y-2 text-xs" style={{ color: C.textSub }}>

<div className="flex justify-between"><span>Latitude</span><b style={{ color: C.text }}>{live.latitude.toFixed(6)}</b></div>

<div className="flex justify-between"><span>Longitude</span><b style={{ color: C.text }}>{live.longitude.toFixed(6)}</b></div>

<div className="flex justify-between"><span>Accuracy</span><b style={{ color: C.text }}>&plusmn; {Math.round(rand(3, 9))} m</b></div>

<div className="flex justify-between"><span>GPS Status</span><Pill color={live.gps_status === 'ONLINE' ? C.green : C.red} bg={live.gps_status === 'ONLINE' ? C.greenBg : C.redBg} border={live.gps_status === 'ONLINE' ? C.greenBorder : C.redBorder}>{live.gps_status}</Pill></div>

          </div>

`<div className="mt-4 pt-4" style={{ borderTop: `1px solid ${C.border}` }}>`

<p className="text-xs font-medium mb-2" style={{ color: C.textSub }}>Recent Locations</p>

<div className="space-y-2 max-h-48 overflow-y-auto pr-1">

{points.slice(0, 6).map((p) => {

const rs = riskStyle(p.risk_level);

return (

<div key={p.id} className="flex items-center justify-between text-[11px]">

<span style={{ color: C.textSub }}>{fmtDateTimeShort(p.timestamp)}</span>

<span style={{ color: C.text }}>{p.latitude.toFixed(3)}, {p.longitude.toFixed(3)}</span>

<span className="w-2 h-2 rounded-full" style={{ backgroundColor: rs.color }} />

                  </div>

);

})}

            </div>

          </div>

        </Card>

      </div>

    </div>

);

}



/* ============================== ALERTS PAGE ============================== */

function AlertsPage({ alerts }) {

const [filter, setFilter] = useState('ALL');

const filtered = filter === 'ALL' ? alerts : alerts.filter((a) => a.type === filter);

return (

<Card>

<SectionTitle action={

<div className="flex gap-1.5 flex-wrap">

{['ALL', 'green', 'yellow', 'orange', 'red'].map((f) => (

<button key={f} onClick={() => setFilter(f)} className="text-[11px] font-medium px-2.5 py-1 rounded-full capitalize" style={{ backgroundColor: filter === f ? C.navy : C.bg, color: filter === f ? '#fff' : C.textSub }}>{f === 'ALL' ? 'All' : f}</button>

))}

        </div>

}>Recent Alerts</SectionTitle>

<div className="space-y-2.5">

{filtered.map((a) => <AlertRow key={a.id} a={a} />)}

{filtered.length === 0 && <p className="text-sm text-center py-8" style={{ color: C.textMute }}>No alerts of this type.</p>}

      </div>

    </Card>

);

}



/* ============================== REPORTS PAGE ============================== */

function ReportsPage({ history }) {

const [from, setFrom] = useState('');

const [to, setTo] = useState('');

const [resultFilter, setResultFilter] = useState('ALL');

const [riskFilter, setRiskFilter] = useState('ALL');

const [status, setStatus] = useState('');



const filteredRows = useMemo(() => {

return history.filter((r) => {

const d = r.timestamp.slice(0, 10);

if (from && d < from) return false;

if (to && d > to) return false;

if (resultFilter !== 'ALL' && r.screening_result !== resultFilter) return false;

if (riskFilter !== 'ALL' && r.risk_level !== riskFilter) return false;

return true;

});

}, [history, from, to, resultFilter, riskFilter]);



const downloadExcel = () => {

const rows = filteredRows.map((r) => ({

'Date': fmtDate(r.timestamp), 'Time': fmtDateTimeShort(r.timestamp), 'Session ID': r.id,

'Temperature (C)': r.temperature, 'AQI': r.aqi, 'Cough Score': r.cough_score,

'Breathing Status': r.breathing_status, 'Face Detected': r.face_detected ? 'Yes' : 'No',

'Screening Result': screeningMeta(r.screening_result).label, 'Confidence (%)': r.confidence,

'Risk Level': r.risk_level, 'Latitude': r.latitude, 'Longitude': r.longitude,

}));

const ws = XLSX.utils.json_to_sheet(rows);

const wb = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(wb, ws, 'Screening Report');

`XLSX.writeFile(wb, `HealthScreen_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);`

setStatus('Excel report downloaded.');

};



const downloadPdf = () => {

const win = window.open('', '_blank');

`const rowsHtml = filteredRows.map((r) => `<tr>`

      <td>${fmtDate(r.timestamp)} ${fmtDateTimeShort(r.timestamp)}</td><td>${r.id}</td>

      <td>${r.temperature.toFixed(1)}</td><td>${r.aqi}</td><td>${r.cough_score.toFixed(2)}</td>

      <td>${r.breathing_status}</td><td>${screeningMeta(r.screening_result).label}</td>

`      <td>${r.confidence}%</td><td>${r.risk_level}</td></tr>`).join('');`

`win.document.write(`<html><head><title>HealthScreen AI Report</title>`

      <style>body{font-family:Arial,sans-serif;padding:24px;color:#0F1B2D}

      h1{font-size:18px;margin-bottom:2px}p{font-size:12px;color:#5B6B84;margin-top:0}

      table{width:100%;border-collapse:collapse;margin-top:16px;font-size:11px}

      th,td{border:1px solid #E3E9F2;padding:6px 8px;text-align:left}

      th{background:#F2F5FA}.foot{margin-top:20px;font-size:10px;color:#8B98AC}</style></head>

      <body><h1>Adaptive Multi-Sensor Health Screening Framework</h1>

      <p>Screening Report &middot; Generated ${new Date().toLocaleString()}</p>

      <table><thead><tr><th>Date/Time</th><th>Session ID</th><th>Temp (\u00b0C)</th><th>AQI</th><th>Cough</th><th>Breathing</th><th>Result</th><th>Confidence</th><th>Risk</th></tr></thead>

      <tbody>${rowsHtml}</tbody></table>

      <p class="foot">This system is an AI-assisted research prototype for non-contact health screening. Results are probabilistic screening indicators and are not a medical diagnosis. Abnormal results should be evaluated by a qualified healthcare professional.</p>

`      </body></html>`);`

win.document.close();

win.focus();

win.print();

setStatus('PDF report opened in a new tab \u2014 use Print > Save as PDF.');

};



return (

<div className="space-y-4">

<Card>

<SectionTitle>Generate Screening Report</SectionTitle>

<div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>From</label>

`<input type="date" value={from} onChange={(e) => setFrom(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }} />`

          </div>

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>To</label>

`<input type="date" value={to} onChange={(e) => setTo(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }} />`

          </div>

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>Screening Result</label>

`<select value={resultFilter} onChange={(e) => setResultFilter(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }}>`

<option value="ALL">All</option>

<option value="NO_ABNORMALITY_DETECTED">No Abnormality Detected</option>

<option value="REQUIRES_FURTHER_SCREENING">Requires Further Screening</option>

<option value="POSSIBLE_ABNORMALITY">Possible Abnormality</option>

<option value="HIGH_RISK_SCREENING">High Risk Screening</option>

            </select>

          </div>

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>Risk Level</label>

`<select value={riskFilter} onChange={(e) => setRiskFilter(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }}>`

<option value="ALL">All</option><option value="LOW">Low</option><option value="MEDIUM">Medium</option><option value="HIGH">High</option>

            </select>

          </div>

        </div>

<p className="text-xs mt-3" style={{ color: C.textSub }}>{filteredRows.length} screening record(s) match this selection.</p>

<div className="flex flex-wrap gap-3 mt-4">

<button onClick={downloadExcel} className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white" style={{ backgroundColor: C.green }}><Download size={15} />Download Excel Report</button>

<button onClick={downloadPdf} className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white" style={{ backgroundColor: C.red }}><Download size={15} />Download PDF Report</button>

`<button onClick={() => setStatus(`Report generated with ${filteredRows.length} records.`)} className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white" style={{ backgroundColor: C.blue }}><FileText size={15} />Generate Screening Report</button>`

        </div>

{status && <p className="text-xs mt-3" style={{ color: C.green }}>{status}</p>}

      </Card>

    </div>

);

}



/* ============================== SETTINGS PAGE ============================== */

function SettingsPage() {

const [rpiIp, setRpiIp] = useState('192.168.1.42');

const [refreshInterval, setRefreshInterval] = useState(3);

const [sensors, setSensors] = useState({ thermal: true, mic: true, aqi: true, gps: true, camera: true });

const [emailAlerts, setEmailAlerts] = useState(true);

const [phoneAlerts, setPhoneAlerts] = useState(false);

const [tempThreshold, setTempThreshold] = useState(37.5);

const [aqiThreshold, setAqiThreshold] = useState(150);

const [saved, setSaved] = useState(false);



return (

<div className="space-y-4">

<Card>

<SectionTitle>System</SectionTitle>

<div className="grid grid-cols-1 sm:grid-cols-2 gap-4">

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>Raspberry Pi IP Address</label>

`<input value={rpiIp} onChange={(e) => setRpiIp(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }} />`

          </div>

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>Data Refresh Interval (seconds)</label>

`<input type="number" min="1" max="30" value={refreshInterval} onChange={(e) => setRefreshInterval(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }} />`

          </div>

<div className="flex items-center justify-between rounded-lg px-3 py-2.5" style={{ backgroundColor: C.bg }}>

<span className="text-sm inline-flex items-center gap-2" style={{ color: C.text }}><Server size={14} />Server / API Status</span>

<Pill color={C.green} bg={C.greenBg} border={C.greenBorder}>Connected</Pill>

          </div>

<div className="flex items-center justify-between rounded-lg px-3 py-2.5" style={{ backgroundColor: C.bg }}>

<span className="text-sm inline-flex items-center gap-2" style={{ color: C.text }}><Camera size={14} />ESP32-CAM Status</span>

<Pill color={C.green} bg={C.greenBg} border={C.greenBorder}>Connected</Pill>

          </div>

        </div>

      </Card>



<Card>

<SectionTitle>Sensors</SectionTitle>

<div className="grid grid-cols-1 sm:grid-cols-2 gap-3">

{[

{ key: 'thermal', label: 'Thermal Sensor', icon: Thermometer },

{ key: 'mic', label: 'Microphone', icon: Mic },

{ key: 'aqi', label: 'AQI Sensor', icon: Wind },

{ key: 'gps', label: 'GPS', icon: MapPin },

{ key: 'camera', label: 'Camera', icon: Camera },

].map((s) => (

<div key={s.key} className="flex items-center justify-between rounded-lg px-3 py-2.5" style={{ backgroundColor: C.bg }}>

<span className="text-sm inline-flex items-center gap-2" style={{ color: C.text }}><s.icon size={14} />{s.label}</span>

<Toggle checked={sensors[s.key]} onChange={(v) => setSensors({ ...sensors, [s.key]: v })} />

            </div>

))}

        </div>

      </Card>



<Card>

<SectionTitle>Alerts</SectionTitle>

<div className="space-y-3">

<div className="flex items-center justify-between rounded-lg px-3 py-2.5" style={{ backgroundColor: C.bg }}>

<span className="text-sm inline-flex items-center gap-2" style={{ color: C.text }}><Mail size={14} />Email Alerts</span>

<Toggle checked={emailAlerts} onChange={setEmailAlerts} />

          </div>

<div className="flex items-center justify-between rounded-lg px-3 py-2.5" style={{ backgroundColor: C.bg }}>

<span className="text-sm inline-flex items-center gap-2" style={{ color: C.text }}><Phone size={14} />Phone Notifications</span>

<Toggle checked={phoneAlerts} onChange={setPhoneAlerts} />

          </div>

<div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2">

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>Temperature Alert Threshold (\u00b0C)</label>

`<input type="number" step="0.1" value={tempThreshold} onChange={(e) => setTempThreshold(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }} />`

            </div>

<div>

<label className="text-xs font-medium" style={{ color: C.textSub }}>AQI Alert Threshold</label>

`<input type="number" value={aqiThreshold} onChange={(e) => setAqiThreshold(e.target.value)} className="w-full mt-1 px-3 py-2 rounded-lg text-sm" style={{ border: `1px solid ${C.border}` }} />`

            </div>

          </div>

        </div>

      </Card>



<div className="flex items-center gap-3">

<button onClick={() => { setSaved(true); setTimeout(() => setSaved(false), 2500); }} className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white" style={{ backgroundColor: C.blue }}>

<Save size={15} />Save Settings

        </button>

{saved && <span className="text-xs" style={{ color: C.green }}>Settings saved for this session.</span>}

      </div>

    </div>

);

}



/* ============================== ABOUT PAGE ============================== */

function AboutPage() {

const tags = ['Raspberry Pi 3B+', 'ESP32-CAM', 'AI / ML', 'IoT', 'Multi-Sensor Data Fusion'];

return (

<div className="space-y-4">

<Card>

<div className="flex items-center gap-3 mb-2">

<div className="w-11 h-11 rounded-xl flex items-center justify-center" style={{ backgroundColor: C.navy }}><Sparkles size={22} color="#fff" /></div>

<div>

<h1 className="text-lg font-semibold" style={{ color: C.text }}>Adaptive Multi-Sensor Health Screening Framework</h1>

<p className="text-sm" style={{ color: C.textSub }}>For Sustainable Urban Communities</p>

          </div>

        </div>

<p className="text-sm mt-3 leading-relaxed" style={{ color: C.textSub }}>

          HealthScreen AI is a non-contact, multi-sensor screening prototype built around a Raspberry Pi 3B+ edge

          controller and an ESP32-CAM vision node. Thermal, audio, air-quality and GPS signals are fused with a

          lightweight AI/ML model to surface probabilistic screening indicators for community health monitoring

          in urban settings.

        </p>

<div className="flex flex-wrap gap-2 mt-4">

{tags.map((t) => <span key={t} className="text-xs font-medium px-3 py-1.5 rounded-full" style={{ backgroundColor: C.blueBg, color: C.blue }}>{t}</span>)}

        </div>

      </Card>



<Card>

<SectionTitle>System Architecture</SectionTitle>

<div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-sm" style={{ color: C.textSub }}>

<div className="flex items-start gap-2"><Cpu size={15} className="mt-0.5" color={C.blue} /><span>Raspberry Pi 3B+ edge controller running the data fusion and AI/ML inference pipeline.</span></div>

<div className="flex items-start gap-2"><Camera size={15} className="mt-0.5" color={C.blue} /><span>ESP32-CAM node for visual and face-presence detection, streamed over Wi-Fi.</span></div>

<div className="flex items-start gap-2"><Thermometer size={15} className="mt-0.5" color={C.blue} /><span>Contactless thermal sensor for body/surface temperature screening.</span></div>

<div className="flex items-start gap-2"><Mic size={15} className="mt-0.5" color={C.blue} /><span>Microphone array for cough and respiratory audio pattern analysis.</span></div>

<div className="flex items-start gap-2"><Wind size={15} className="mt-0.5" color={C.blue} /><span>Environmental AQI sensor for local air-quality context.</span></div>

<div className="flex items-start gap-2"><MapPin size={15} className="mt-0.5" color={C.blue} /><span>GPS module for geotagging screening sessions across urban sites.</span></div>

        </div>

      </Card>



<Card style={{ backgroundColor: C.yellowBg, borderColor: C.yellowBorder }}>

<div className="flex items-start gap-2.5">

<ShieldCheck size={18} color={C.yellow} className="flex-shrink-0 mt-0.5" />

<p className="text-xs leading-relaxed" style={{ color: C.yellow }}>

            This system is an AI-assisted research prototype for non-contact health screening. Results are

            probabilistic screening indicators and are not a medical diagnosis. Abnormal results should be

            evaluated by a qualified healthcare professional.

          </p>

        </div>

      </Card>

    </div>

);

}



/* ============================== FOOTER ============================== */

function Footer() {

return (

<div className="text-center text-[11px] py-4 px-4" style={{ color: C.textMute }}>

      This system is an AI-assisted research prototype for non-contact health screening. Results are probabilistic

      screening indicators and are not a medical diagnosis. Abnormal results should be evaluated by a qualified

      healthcare professional.

    </div>

);

}



/* ============================== APP ============================== */

export default function App() {

const [page, setPage] = useState('dashboard');

const [mobileOpen, setMobileOpen] = useState(false);

const [demoMode, setDemoMode] = useState(true);

const [now, setNow] = useState(new Date());

const [live, setLive] = useState(() => genReading());

const [series, setSeries] = useState(() => seedHistory(60).reverse());

const [history, setHistory] = useState(() => seedHistory(48));

const [alerts, setAlerts] = useState(() => seedHistory(6).map(alertFromReading));

const [chartRange, setChartRange] = useState('10m');



useEffect(() => {

const t = setInterval(() => setNow(new Date()), 1000);

return () => clearInterval(t);

}, []);



useEffect(() => {

if (!demoMode) return;

const t = setInterval(() => {

const r = genReading();

setLive(r);

setSeries((s) => [...s.slice(-119), r]);

setHistory((h) => [r, ...h].slice(0, 300));

const a = alertFromReading(r);

if (a.type !== 'green' || Math.random() > 0.6) {

setAlerts((al) => [a, ...al].slice(0, 60));

}

}, 3000);

return () => clearInterval(t);

}, [demoMode]);



const highAlertCount = alerts.filter((a) => a.type === 'red' || a.type === 'orange').length;



return (

<div className="min-h-screen w-full flex" style={{ backgroundColor: C.bg, fontFamily: 'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif' }}>

<Sidebar page={page} setPage={setPage} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} alertCount={highAlertCount} />

<div className="flex-1 flex flex-col min-w-0 lg:ml-64">

<Header setMobileOpen={setMobileOpen} live={live} now={now} demoMode={demoMode} setDemoMode={setDemoMode} systemOnline={true} rpiOnline={live.thermal_status === 'ONLINE'} />

<main className="flex-1 p-4 sm:p-6">

{page === 'dashboard' && <DashboardPage live={live} series={series} alerts={alerts} chartRange={chartRange} setChartRange={setChartRange} />}

{page === 'live' && <LiveMonitoringPage live={live} />}

{page === 'history' && <HistoryPage history={history} />}

{page === 'analytics' && <AnalyticsPage history={history} />}

{page === 'map' && <MapPage live={live} history={history} />}

{page === 'alerts' && <AlertsPage alerts={alerts} />}

{page === 'reports' && <ReportsPage history={history} />}

{page === 'settings' && <SettingsPage />}

{page === 'about' && <AboutPage />}

        </main>

<Footer />

      </div>

    </div>

);

}
