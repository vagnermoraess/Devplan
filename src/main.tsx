import React, {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { createRoot } from "react-dom/client";
import ExcelJS from "exceljs";
import * as I from "lucide-react";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import "./styles.css";
import "./dashboard.css";
import "./capacity.css";
import "./changes.css";
import "./demands.css";

type Page =
  | "Visão geral"
  | "Roadmap"
  | "Demandas"
  | "Produtos"
  | "Capacidade"
  | "Planejamento"
  | "Riscos"
  | "Mudanças"
  | "Relatórios"
  | "Configurações"
  | "Usuários";
type Demand = {
  id: string;
  name: string;
  product: string;
  owner: string;
  status: string;
  priority: string;
  effort: number;
  progress: number;
  due: string;
  risk: string;
  type?: string;
  version?: string;
  notes?: string;
  startDate?: string;
  dueDate?: string;
  resourceIds?: number[];
};
type ChangeLog={id:number;date:string;title:string;actor:string;detail:string;tone:"blue"|"good"|"warn"|"bad";product?:string;quarter?:string};
type Product = {
  id: number;
  name: string;
  code: string;
  description: string;
  team: string;
  coordinator: string;
  active: boolean;
};
const initialProducts: Product[] = [
  {
    id: 1,
    name: "Commerce",
    code: "COM",
    description: "Experiência de comércio digital",
    team: "Squad Commerce",
    coordinator: "Marina Costa",
    active: true,
  },
  {
    id: 2,
    name: "Payments",
    code: "PAY",
    description: "Pagamentos e serviços financeiros",
    team: "Squad Payments",
    coordinator: "Rafael Lima",
    active: true,
  },
  {
    id: 3,
    name: "Customer",
    code: "CUS",
    description: "Jornada e relacionamento com clientes",
    team: "Squad Customer",
    coordinator: "Camila Souza",
    active: true,
  },
  {
    id: 4,
    name: "Platform",
    code: "PLT",
    description: "Plataforma e serviços compartilhados",
    team: "Squad Platform",
    coordinator: "Bruno Alves",
    active: true,
  },
  {
    id: 5,
    name: "Analytics",
    code: "ANA",
    description: "Dados e inteligência analítica",
    team: "Squad Analytics",
    coordinator: "Bianca Melo",
    active: true,
  },
];
const ProductContext = createContext<{
  products: Product[];
  setProducts: React.Dispatch<React.SetStateAction<Product[]>>;
} | null>(null);
function ProductProvider({ children }: { children: React.ReactNode }) {
  const [products, setProducts] = useState(initialProducts);
  return (
    <ProductContext.Provider value={{ products, setProducts }}>
      {children}
    </ProductContext.Provider>
  );
}
function useProducts() {
  const value = useContext(ProductContext);
  if (!value) throw new Error("ProductProvider ausente");
  return value;
}
type RoadmapEntry = {
  demandId: string;
  quarter: string;
  collaborators: string[];
  allocations: { staffId: number; hours: number }[];
};
const demands: Demand[] = [
  {
    id: "DEV-142",
    name: "Novo checkout omnichannel",
    product: "Commerce",
    owner: "Marina",
    status: "Desenvolvimento",
    priority: "Crítica",
    effort: 120,
    progress: 68,
    due: "12 Set",
    startDate: "2026-07-08", dueDate: "2026-09-12",
    risk: "Alto",
  },
  {
    id: "DEV-148",
    name: "Motor antifraude v2",
    product: "Payments",
    owner: "Rafael",
    status: "Em testes",
    priority: "Alta",
    effort: 96,
    progress: 82,
    due: "06 Set",
    startDate: "2026-07-22", dueDate: "2026-09-06",
    risk: "Médio",
  },
  {
    id: "DEV-151",
    name: "Portal de autoatendimento",
    product: "Customer",
    owner: "Camila",
    status: "Planejada",
    priority: "Alta",
    effort: 140,
    progress: 24,
    due: "24 Set",
    startDate: "2026-08-03", dueDate: "2026-09-24",
    risk: "Crítico",
  },
  {
    id: "DEV-154",
    name: "Adequação regulatória BACEN",
    product: "Payments",
    owner: "Lucas",
    status: "Em análise",
    priority: "Crítica",
    effort: 60,
    progress: 10,
    due: "18 Set",
    startDate: "2026-08-24", dueDate: "2026-09-18",
    risk: "Alto",
  },
  {
    id: "DEV-156",
    name: "Otimização de busca",
    product: "Platform",
    owner: "Bruno",
    status: "Concluída",
    priority: "Média",
    effort: 48,
    progress: 100,
    due: "28 Ago",
    startDate: "2026-07-15", dueDate: "2026-08-28",
    risk: "Baixo",
  },
  {
    id: "DEV-159",
    name: "Novo painel de clientes",
    product: "Analytics",
    owner: "Bianca",
    status: "Bloqueada",
    priority: "Média",
    effort: 72,
    progress: 35,
    due: "30 Set",
    startDate: "2026-09-01", dueDate: "2026-09-30",
    risk: "Alto",
  },
];
const DemandContext = createContext<{
  rows: Demand[];
  setRows: React.Dispatch<React.SetStateAction<Demand[]>>;
  roadmap: RoadmapEntry[];
  setRoadmap: React.Dispatch<React.SetStateAction<RoadmapEntry[]>>;
  history: ChangeLog[];
  setHistory: React.Dispatch<React.SetStateAction<ChangeLog[]>>;
} | null>(null);
function DemandProvider({ children }: { children: React.ReactNode }) {
  const [rows, setRows] = useState<Demand[]>(demands);
  const [roadmap, setRoadmap] = useState<RoadmapEntry[]>(
    demands.slice(0, 5).map((d) => ({
      demandId: d.id,
      quarter: "Q3 2026",
      collaborators: [d.owner],
      allocations: [{ staffId: Math.max(1, members.findIndex(([name]) => String(name).startsWith(d.owner)) + 1), hours: d.effort }],
    })),
  );
  const [history,setHistory]=useState<ChangeLog[]>([]);
  return (
    <DemandContext.Provider value={{ rows, setRows, roadmap, setRoadmap, history, setHistory }}>
      {children}
    </DemandContext.Provider>
  );
}
function useDemands() {
  const value = useContext(DemandContext);
  if (!value) throw new Error("DemandProvider ausente");
  return value;
}
const members = [
  ["Marina Costa", "Tech Lead", 130, 118],
  ["Rafael Lima", "Backend", 130, 125],
  ["Camila Souza", "Frontend", 130, 102],
  ["Lucas Rocha", "Backend", 120, 111],
  ["Bruno Alves", "Full Stack", 130, 88],
  ["Bianca Melo", "UX Engineer", 110, 76],
];
const nav: [Page, any][] = [
  ["Visão geral", I.LayoutDashboard],
  ["Demandas", I.ListTodo],
  ["Roadmap", I.Map],
  ["Produtos", I.Boxes],
  ["Capacidade", I.Users],
  ["Planejamento", I.CalendarRange],
  ["Riscos", I.ShieldAlert],
  ["Mudanças", I.GitCompareArrows],
  ["Relatórios", I.BarChart3],
  ["Configurações", I.Settings],
  ["Usuários", I.UserCog],
];
const cap = [
  { m: "Jul", roadmap: 220, sustentacao: 42, novas: 18, reserva: 28 },
  { m: "Ago", roadmap: 248, sustentacao: 48, novas: 34, reserva: 22 },
  { m: "Set", roadmap: 252, sustentacao: 50, novas: 28, reserva: 50 },
];
const trend = [
  { w: "S1", plan: 22, done: 18 },
  { w: "S2", plan: 34, done: 28 },
  { w: "S3", plan: 48, done: 42 },
  { w: "S4", plan: 62, done: 56 },
  { w: "S5", plan: 76, done: 66 },
  { w: "S6", plan: 88, done: 78 },
];
const C = {
  blue: "#246bfd",
  cyan: "#38bdf8",
  green: "#16a36a",
  amber: "#f59e0b",
  red: "#ef4444",
  purple: "#8b5cf6",
};

function Badge({
  children,
  tone = "blue",
}: {
  children: React.ReactNode;
  tone?: string;
}) {
  return <span className={"badge " + tone}>{children}</span>;
}
function Card({
  children,
  className = "",
}: {
  children: React.ReactNode;
  className?: string;
}) {
  return <section className={"card " + className}>{children}</section>;
}
function PageHead({
  title,
  desc,
  action,
}: {
  title: string;
  desc: string;
  action?: React.ReactNode;
}) {
  return (
    <div className="pagehead">
      <div>
        <div className="eyebrow">Q3 2026 · PLANEJAMENTO ATIVO</div>
        <h1>{title}</h1>
        <p>{desc}</p>
      </div>
      {action}
    </div>
  );
}
function ChartTip({ active, payload, label }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div className="charttip">
      <b>{label}</b>
      {payload.map((p: any) => (
        <div key={p.name}>
          <i style={{ background: p.color }} /> {p.name}: <b>{p.value}h</b>
        </div>
      ))}
    </div>
  );
}

function Dashboard({ go }: { go: (p: Page) => void }) {
  const { rows, roadmap } = useDemands();
  const { products } = useProducts();
  const { team } = useTeam();
  const [productFilter, setProductFilter] = useState("Todos");
  const scoped = productFilter === "Todos" ? rows : rows.filter((d) => d.product === productFilter);
  const scopedIds = new Set(scoped.map((d) => d.id));
  const planned = roadmap.filter((item) => scopedIds.has(item.demandId));
  const effort = scoped.reduce((sum, demand) => sum + demand.effort, 0);
  const scopedTeam = productFilter === "Todos" ? team : team.filter((person) => person.product === productFilter);
  const totalCapacity = scopedTeam.reduce((sum, person) => sum + person.total, 0);
  const completed = scoped.filter((d) => d.status === "Concluída").length;
  const critical = scoped.filter((d) => d.risk === "Crítico" || d.risk === "Alto").length;
  const progress = scoped.length ? Math.round(scoped.reduce((sum, d) => sum + d.progress, 0) / scoped.length) : 0;
  const kpis = [
    ["Capacidade total", `${totalCapacity}h`, `${scopedTeam.length} colaboradores`, "blue", I.Users],
    ["Capacidade estimada", `${effort}h`, `${scoped.length} demandas`, "warn", I.Gauge],
    ["Roadmap planejado", String(planned.length), `${completed} concluídas`, "blue", I.Map],
    ["Progresso médio", `${progress}%`, "no escopo atual", "good", I.Crosshair],
    ["Riscos altos/críticos", String(critical), productFilter === "Todos" ? "visão geral" : productFilter, "bad", I.ShieldAlert],
  ];
  return (
    <>
      <PageHead
        title="Visão geral do roadmap"
        desc="Acompanhe capacidade, progresso e decisões críticas do trimestre."
        action={
          <button className="primary" onClick={() => go("Planejamento")}>
            <I.Sparkles /> Simular mudança
          </button>
        }
      />
      <div className="dashboard-filter"><div><I.Filter/><span><b>Escopo dos indicadores</b><small>Visualize o consolidado ou um produto específico</small></span></div><select value={productFilter} onChange={(e) => setProductFilter(e.target.value)}><option>Todos</option>{products.filter((p)=>p.active).map((p)=><option key={p.id} value={p.name}>{p.name}</option>)}</select></div>
      <div className="kpis dashboard-kpis">
        {kpis.map(([n, v, d, t, Icon]: any) => (
          <Card key={n} className="kpi">
            <div className={"kicon " + t}>
              <Icon />
            </div>
            <div className="klabel">
              {n}
              <I.Info />
            </div>
            <strong>{v}</strong>
            <small className={t}>
              {t === "bad" ? "↑" : "↗"} {d} <span>vs. Q2</span>
            </small>
          </Card>
        ))}
      </div>
      <div className="grid2">
        <Card>
          <div className="cardhead">
            <div>
              <h3>Capacidade por categoria</h3>
              <p>Distribuição mensal do trimestre</p>
            </div>
            <button className="icon">
              <I.MoreHorizontal />
            </button>
          </div>
          <div className="chart">
            <ResponsiveContainer>
              <BarChart data={cap} barGap={0}>
                <CartesianGrid vertical={false} />
                <XAxis dataKey="m" />
                <YAxis />
                <Tooltip content={<ChartTip />} />
                <Bar
                  dataKey="roadmap"
                  name="Roadmap"
                  stackId="a"
                  fill={C.blue}
                />
                <Bar
                  dataKey="sustentacao"
                  name="Sustentação"
                  stackId="a"
                  fill={C.cyan}
                />
                <Bar
                  dataKey="novas"
                  name="Novas demandas"
                  stackId="a"
                  fill={C.amber}
                />
                <Bar
                  dataKey="reserva"
                  name="Reserva"
                  stackId="a"
                  fill="#dbe5f4"
                  radius={[5, 5, 0, 0]}
                />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="legend">
            <span>
              <i className="b" />
              Roadmap 69%
            </span>
            <span>
              <i className="c" />
              Sustentação 13%
            </span>
            <span>
              <i className="a" />
              Novas 8%
            </span>
            <span>
              <i />
              Reserva 10%
            </span>
          </div>
        </Card>
        <Card>
          <div className="cardhead">
            <div>
              <h3>Progresso do roadmap</h3>
              <p>Planejado vs. realizado</p>
            </div>
            <Badge tone="good">+6,4%</Badge>
          </div>
          <div className="chart">
            <ResponsiveContainer>
              <AreaChart data={trend}>
                <defs>
                  <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
                    <stop stopColor={C.blue} stopOpacity=".22" />
                    <stop offset="1" stopColor={C.blue} stopOpacity="0" />
                  </linearGradient>
                </defs>
                <CartesianGrid vertical={false} />
                <XAxis dataKey="w" />
                <YAxis />
                <Tooltip />
                <Area
                  type="monotone"
                  dataKey="plan"
                  stroke="#a9b5c7"
                  fill="transparent"
                  strokeDasharray="4 4"
                />
                <Area
                  type="monotone"
                  dataKey="done"
                  stroke={C.blue}
                  fill="url(#g)"
                  strokeWidth={3}
                />
              </AreaChart>
            </ResponsiveContainer>
          </div>
          <div className="progressfoot">
            <b>14 de 18 demandas no prazo</b>
            <span>78% concluído ou em curso</span>
          </div>
        </Card>
      </div>
      <div className="grid3">
        <Card className="span2">
          <div className="cardhead">
            <div>
              <h3>Demandas que exigem atenção</h3>
              <p>Priorizadas por risco e proximidade do prazo</p>
            </div>
            <button className="ghost" onClick={() => go("Demandas")}>
              Ver todas <I.ArrowRight />
            </button>
          </div>
          <DemandTable compact />
        </Card>
        <Card>
          <div className="cardhead">
            <div>
              <h3>Saúde do trimestre</h3>
              <p>Visão consolidada</p>
            </div>
          </div>
          <div className="donut">
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={[{ v: 78 }, { v: 22 }]}
                  dataKey="v"
                  innerRadius={62}
                  outerRadius={78}
                  startAngle={90}
                  endAngle={-270}
                >
                  <Cell fill={C.green} />
                  <Cell fill="#e9eef5" />
                </Pie>
              </PieChart>
            </ResponsiveContainer>
            <div>
              <strong>78</strong>
              <span>/100</span>
              <small>SAUDÁVEL</small>
            </div>
          </div>
          <div className="health">
            <p>
              <span>
                <i className="good" />
                Prazo
              </span>
              <b>82%</b>
            </p>
            <p>
              <span>
                <i className="warn" />
                Capacidade
              </span>
              <b>91%</b>
            </p>
            <p>
              <span>
                <i className="bad" />
                Risco
              </span>
              <b>3 críticos</b>
            </p>
          </div>
        </Card>
      </div>
    </>
  );
}

function DemandTable({ compact = false }: { compact?: boolean }) {
  const rows = compact ? demands.slice(0, 4) : demands;
  return (
    <div className="tablewrap">
      <table>
        <thead>
          <tr>
            <th>Demanda</th>
            <th>Produto</th>
            <th>Responsável</th>
            <th>Status</th>
            <th>Prioridade</th>
            <th>Esforço</th>
            <th>Prazo</th>
            <th>Risco</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((d) => (
            <tr key={d.id}>
              <td>
                <b>{d.name}</b>
                <small>{d.id}</small>
              </td>
              <td>{d.product}</td>
              <td>
                <span className="person">{d.owner[0]}</span>
                {d.owner}
              </td>
              <td>
                <Badge
                  tone={
                    d.status === "Concluída"
                      ? "good"
                      : d.status === "Bloqueada"
                        ? "bad"
                        : "blue"
                  }
                >
                  {d.status}
                </Badge>
              </td>
              <td>
                <span className={"priority " + d.priority.toLowerCase()}>
                  <i />
                  {d.priority}
                </span>
              </td>
              <td>{d.effort}h</td>
              <td>{d.due}</td>
              <td>
                <Badge
                  tone={
                    d.risk === "Crítico"
                      ? "bad"
                      : d.risk === "Alto"
                        ? "warn"
                        : d.risk === "Baixo"
                          ? "good"
                          : "gray"
                  }
                >
                  {d.risk}
                </Badge>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const quarterMonths: Record<string, string[]> = {
  Q1: ["JANEIRO", "FEVEREIRO", "MARÇO"],
  Q2: ["ABRIL", "MAIO", "JUNHO"],
  Q3: ["JULHO", "AGOSTO", "SETEMBRO"],
  Q4: ["OUTUBRO", "NOVEMBRO", "DEZEMBRO"],
};
const quarters = [
  "Q1 2026",
  "Q2 2026",
  "Q3 2026",
  "Q4 2026",
  "Q1 2027",
  "Q2 2027",
  "Q3 2027",
  "Q4 2027",
];

function Roadmap() {
  const { rows, roadmap, setRoadmap, setHistory } = useDemands();
  const { team } = useTeam();
  const { products } = useProducts();
  const [view, setView] = useState("Timeline");
  const [quarter, setQuarter] = useState("Q3 2026");
  const [productFilter, setProductFilter] = useState("Todos");
  const [picker, setPicker] = useState(false);
  const [staffFor, setStaffFor] = useState<RoadmapEntry | null>(null);
  const items = roadmap.flatMap((entry) => {
    const demand = rows.find((d) => d.id === entry.demandId);
    return entry.quarter === quarter && demand && (productFilter === "Todos" || demand.product === productFilter) ? [{ demand, entry }] : [];
  });
  const available = rows.filter(
    (d) => !roadmap.some((r) => r.demandId === d.id),
  );
  const months = quarterMonths[quarter.slice(0, 2)];
  const register=(title:string,detail:string,tone:ChangeLog["tone"]="blue",eventQuarter?:string)=>{const demandId=detail.match(/DEV-\d+/)?.[0];const product=rows.find((d)=>d.id===demandId)?.product;setHistory((current)=>[{id:Date.now(),date:new Date().toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"}),title,actor:"Vagner Moraes",detail,tone,product,quarter:eventQuarter},...current])};
  const add = (id: string) => {
    setRoadmap((current) => [
      ...current,
      { demandId: id, quarter, collaborators: [], allocations: [] },
    ]);
    register("Demanda adicionada ao roadmap",`${id} planejada em ${quarter}`,"good",quarter);
    setPicker(false);
  };
  const move = (id: string, destination: string) => {
    const origin=roadmap.find((item)=>item.demandId===id)?.quarter;
    setRoadmap((current) =>
      current.map((r) =>
        r.demandId === id ? { ...r, quarter: destination } : r,
      ),
    );
    register("Demanda movida de quarter",`${id}: ${origin} → ${destination}`,"warn",destination);
  };
  const remove = (id: string) => {const origin=roadmap.find((item)=>item.demandId===id)?.quarter;setRoadmap((current) => current.filter((r) => r.demandId !== id));register("Demanda removida do roadmap",`${id} removida do planejamento`,"bad",origin)};
  const updateAllocation = (person: Staff, hours: number) => {
    if (!staffFor) return;
    const allocations = hours > 0
      ? [...staffFor.allocations.filter((a) => a.staffId !== person.id), { staffId: person.id, hours }]
      : staffFor.allocations.filter((a) => a.staffId !== person.id);
    const collaborators = allocations.map((a) => team.find((member) => member.id === a.staffId)?.name).filter(Boolean) as string[];
    const updated = { ...staffFor, collaborators, allocations };
    setStaffFor(updated);
    setRoadmap((current) =>
      current.map((r) => (r.demandId === updated.demandId ? updated : r)),
    );
    register("Alocação do roadmap alterada",`${updated.demandId}: ${allocations.reduce((sum,a)=>sum+a.hours,0)}h distribuídas entre ${collaborators.length} colaborador(es)`,"blue",updated.quarter);
  };
  return (
    <>
      <PageHead
        title="Roadmap trimestral"
        desc="Planeje demandas por quarter e defina os colaboradores responsáveis."
        action={
          <button className="primary" onClick={() => setPicker(true)}>
            <I.ListPlus /> Planejar demanda cadastrada
          </button>
        }
      />
      <div className="roadmap-source">
        <I.Link2 />
        <span>
          <b>Demandas são a fonte única.</b> Não é possível criar itens
          diretamente no Roadmap; aqui você apenas planeja demandas já
          cadastradas.
        </span>
        <Badge tone="good">{items.length} neste quarter</Badge>
      </div>
      <div className="toolbar">
        <div className="seg">
          {["Timeline", "Lista"].map((x) => (
            <button
              key={x}
              className={view === x ? "active" : ""}
              onClick={() => setView(x)}
            >
              {x}
            </button>
          ))}
        </div>
        <label className="quarter-select">
          <I.Boxes />
          <select value={productFilter} onChange={(e) => setProductFilter(e.target.value)}>
            <option>Todos</option>
            {products.filter((product) => product.active).map((product) => <option key={product.id} value={product.name}>{product.name}</option>)}
          </select>
        </label>
        <label className="quarter-select">
          <I.CalendarRange />
          <select value={quarter} onChange={(e) => setQuarter(e.target.value)}>
            {quarters.map((q) => (
              <option key={q}>{q}</option>
            ))}
          </select>
        </label>
      </div>
      {items.length === 0 ? (
        <Card className="roadmap-empty">
          <I.Map />
          <h3>Nenhuma demanda em {quarter}</h3>
          <p>
            Planeje uma demanda cadastrada ou mova uma demanda de outro quarter.
          </p>
          <button className="primary" onClick={() => setPicker(true)}>
            <I.ListPlus /> Selecionar demanda
          </button>
        </Card>
      ) : view === "Timeline" ? (
        <Timeline
          items={items}
          months={months}
          quarter={quarter}
          onMove={move}
          onStaff={setStaffFor}
        />
      ) : (
        <RoadmapList
          items={items}
          onMove={move}
          onStaff={setStaffFor}
          onRemove={remove}
        />
      )}
      {picker && (
        <div
          className="overlay roadmap-picker-overlay"
          onMouseDown={() => setPicker(false)}
        >
          <div
            className="roadmap-picker"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-head">
              <div>
                <span className="modal-icon">
                  <I.ListPlus />
                </span>
                <div>
                  <h2>Planejar demanda em {quarter}</h2>
                  <p>
                    Somente demandas previamente cadastradas estão disponíveis.
                  </p>
                </div>
              </div>
              <button className="modal-close" onClick={() => setPicker(false)}>
                <I.X />
              </button>
            </div>
            <div className="picker-list">
              {available.length === 0 ? (
                <div className="picker-empty">
                  <I.CircleCheck />
                  <b>Todas as demandas já estão planejadas</b>
                  <span>
                    Cadastre uma nova demanda para disponibilizá-la aqui.
                  </span>
                </div>
              ) : (
                available.map((d) => (
                  <button key={d.id} onClick={() => add(d.id)}>
                    <span className="person">{d.owner[0]}</span>
                    <div>
                      <b>{d.name}</b>
                      <small>
                        {d.id} · {d.product} · {d.effort}h
                      </small>
                    </div>
                    <Badge tone={d.priority === "Crítica" ? "bad" : "gray"}>
                      {d.priority}
                    </Badge>
                    <I.Plus />
                  </button>
                ))
              )}
            </div>
            <div className="picker-foot">
              <button className="ghost" onClick={() => setPicker(false)}>
                Cancelar
              </button>
            </div>
          </div>
        </div>
      )}
      {staffFor && (
        <div
          className="overlay confirm-overlay"
          onMouseDown={() => setStaffFor(null)}
        >
          <div
            className="staff-roadmap-modal"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="modal-head">
              <div>
                <span className="modal-icon">
                  <I.UsersRound />
                </span>
                <div>
                  <h2>Colaboradores da demanda</h2>
                  <p>{rows.find((d) => d.id === staffFor.demandId)?.name}</p>
                </div>
              </div>
              <button className="modal-close" onClick={() => setStaffFor(null)}>
                <I.X />
              </button>
            </div>
            <div className="staff-check-list">
              {team.filter((person) => person.product === rows.find((d) => d.id === staffFor.demandId)?.product).map((person) => (
                <label key={person.id}>
                  <span className="avatar">
                    {person.name
                      .split(" ")
                      .map((x) => x[0])
                      .slice(0, 2)
                      .join("")}
                  </span>
                  <div>
                    <b>{person.name}</b>
                    <small>
                      {person.role} · {person.total - person.used}h disponíveis
                    </small>
                  </div>
                  <div className="allocation-hours"><input type="number" min="0" max={person.total-person.used} value={staffFor.allocations.find((a) => a.staffId === person.id)?.hours || 0} onChange={(e) => updateAllocation(person, Number(e.target.value))}/><span>horas</span></div>
                </label>
              ))}
            </div>
            <div className="allocation-total"><span>Total alocado</span><b>{staffFor.allocations.reduce((sum,a)=>sum+a.hours,0)}h / {rows.find((d)=>d.id===staffFor.demandId)?.effort || 0}h</b></div>
            <div className="picker-foot">
              <button className="primary" onClick={() => setStaffFor(null)}>
                <I.Check /> Concluir
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
function Timeline({
  items,
  months,
  quarter,
  onMove,
  onStaff,
}: {
  items: { demand: Demand; entry: RoadmapEntry }[];
  months: string[];
  quarter: string;
  onMove: (id: string, quarter: string) => void;
  onStaff: (entry: RoadmapEntry) => void;
}) {
  const [quarterName,yearText]=quarter.split(" ");
  const quarterIndex=Number(quarterName.slice(1))-1;
  const quarterStart=new Date(Number(yearText),quarterIndex*3,1).getTime();
  const quarterEnd=new Date(Number(yearText),quarterIndex*3+3,0,23,59,59).getTime();
  const span=quarterEnd-quarterStart;
  const position=(d:Demand)=>{const start=Math.max(quarterStart,new Date(`${d.startDate || `${yearText}-${String(quarterIndex*3+1).padStart(2,"0")}-01`}T12:00:00`).getTime());const end=Math.min(quarterEnd,new Date(`${d.dueDate || d.startDate || `${yearText}-${String(quarterIndex*3+3).padStart(2,"0")}-28`}T12:00:00`).getTime());const left=Math.max(0,Math.min(100,(start-quarterStart)/span*100));const width=Math.max(4,Math.min(100-left,(end-start)/span*100));return {left:`${left}%`,width:`${width}%`}};
  return (
    <Card>
      <div className="timeline">
        <div className="tlhead">
          <b>Demanda</b>
          {months.map((month) => (
            <span key={month}>{month}</span>
          ))}
        </div>
        {items.map(({ demand: d, entry }, i) => (
          <div className="tlrow" key={d.id}>
            <div>
              <b>{d.name}</b>
              <small>
                {d.id} · {entry.collaborators.length || 0} colaborador(es) ·{" "}
                {d.effort}h
              </small>
            </div>
            <div className="track">
              <span
                className={"bar b" + (i % 5)}
                style={position(d)}
                onClick={() => onStaff(entry)}
                title="Ver colaboradores alocados"
              >
                <em>{d.progress}% · {d.startDate ? new Date(d.startDate+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"short"}) : "início não informado"}</em>
              </span>
              <i style={{ left: "66%" }} />
              <div className="timeline-actions">
                <button
                  title="Definir colaboradores"
                  onClick={() => onStaff(entry)}
                >
                  <I.Users />
                </button>
                <select
                  title="Mover para outro quarter"
                  value={entry.quarter}
                  onChange={(e) => onMove(d.id, e.target.value)}
                >
                  {quarters.map((q) => (
                    <option key={q}>{q}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        ))}
      </div>
    </Card>
  );
}
function RoadmapList({
  items,
  onMove,
  onStaff,
  onRemove,
}: {
  items: { demand: Demand; entry: RoadmapEntry }[];
  onMove: (id: string, quarter: string) => void;
  onStaff: (entry: RoadmapEntry) => void;
  onRemove: (id: string) => void;
}) {
  return (
    <Card>
      <div className="tablewrap">
        <table>
          <thead>
            <tr>
              <th>Demanda</th>
              <th>Produto</th>
              <th>Colaboradores</th>
              <th>Status</th>
              <th>Prioridade</th>
              <th>Esforço</th>
              <th>Quarter</th>
              <th>Ação</th>
            </tr>
          </thead>
          <tbody>
            {items.map(({ demand: d, entry }) => (
              <tr key={d.id}>
                <td>
                  <b>{d.name}</b>
                  <small>{d.id}</small>
                </td>
                <td>{d.product}</td>
                <td>
                  <button className="staff-link" onClick={() => onStaff(entry)}>
                    <I.Users />
                    {entry.collaborators.length
                      ? `${entry.collaborators.length} alocados`
                      : "Adicionar"}
                  </button>
                </td>
                <td>
                  <Badge>{d.status}</Badge>
                </td>
                <td>{d.priority}</td>
                <td>{d.effort}h</td>
                <td>
                  <select
                    className="quarter-inline"
                    value={entry.quarter}
                    onChange={(e) => onMove(d.id, e.target.value)}
                  >
                    {quarters.map((q) => (
                      <option key={q}>{q}</option>
                    ))}
                  </select>
                </td>
                <td>
                  <button
                    className="unlink-button"
                    onClick={() => onRemove(d.id)}
                  >
                    <I.Unlink /> Remover
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

const emptyDemandForm = {
  name: "",
  description: "",
  product: "Commerce",
  type: "Evolutiva",
  version: "",
  origin: "Produto",
  requester: "",
  priority: "Média",
  status: "Backlog",
  due: "",
  startDate: "",
  effort: "40",
  owner: "",
  dependencies: "",
  justification: "",
  impact: "",
  resourceIds: [] as number[],
  planningQuarter: "Planejamento",
};
function Demands() {
  const [q, setQ] = useState("");
  const { rows, setRows, roadmap, setRoadmap, setHistory } = useDemands();
  const { products } = useProducts();
  const { team } = useTeam();
  const [modal, setModal] = useState(false);
  const [saved, setSaved] = useState(false);
  const [editing, setEditing] = useState<Demand | null>(null);
  const [removing, setRemoving] = useState<Demand | null>(null);
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [form, setForm] = useState(emptyDemandForm);
  const fileRef=useRef<HTMLInputElement>(null);
  const [importing,setImporting]=useState(false);
  const [importFeedback,setImportFeedback]=useState("");
  const [productFilter,setProductFilter]=useState("Todos");
  const [quarterFilter,setQuarterFilter]=useState("Todos");
  const selectedProduct = products.find((product) => product.name === form.product);
  const availableResources = team.filter((person) => person.product === form.product);
  const filtered = rows.filter((d) => {
    const planning=roadmap.find((item)=>item.demandId===d.id);
    const matchesQuarter=quarterFilter==="Todos"||(quarterFilter==="Planejamento"?!planning:planning?.quarter===quarterFilter);
    return (d.name+d.id+d.product).toLowerCase().includes(q.toLowerCase())&&(productFilter==="Todos"||d.product===productFilter)&&matchesQuarter;
  });
  const change = (
    e: React.ChangeEvent<
      HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement
    >,
  ) => setForm(e.target.name === "product" ? { ...form, product: e.target.value, resourceIds: [] } : { ...form, [e.target.name]: e.target.value });
  const importExcel=async(e:React.ChangeEvent<HTMLInputElement>)=>{const file=e.target.files?.[0];if(!file)return;setImporting(true);setImportFeedback("");try{const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(await file.arrayBuffer());const sheet=workbook.worksheets[0];if(!sheet)throw new Error("Planilha sem conteúdo");const headers:Record<number,string>={};sheet.getRow(1).eachCell((cell,col)=>{headers[col]=String(cell.text).trim().toLowerCase()});const imported:Demand[]=[];let ignored=0;const text=(row:ExcelJS.Row,names:string[])=>{const col=Object.entries(headers).find(([,header])=>names.includes(header))?.[0];return col?row.getCell(Number(col)).text.trim():""};const iso=(value:string)=>{if(!value)return "";const match=value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);return match?`${match[3]}-${match[2].padStart(2,"0")}-${match[1].padStart(2,"0")}`: /^\d{4}-\d{2}-\d{2}$/.test(value)?value:""};const nextId=Math.max(...rows.map((d)=>Number(d.id.replace(/\D/g,""))||0),159)+1;sheet.eachRow((row,rowNumber)=>{if(rowNumber===1)return;const name=text(row,["demanda","título","titulo"]);const productText=text(row,["produto"]);const product=products.find((p)=>p.name.toLowerCase()===productText.toLowerCase()&&p.active);if(!name||!product){ignored++;return}const dueDate=iso(text(row,["prazo","prazo desejado","data fim","data de fim"]));const effort=Number(text(row,["capacidade estimada","esforço","esforco","horas"]).replace(",","."))||0;imported.push({id:`DEV-${nextId+imported.length}`,name,product:product.name,owner:product.coordinator.split(" ")[0],status:text(row,["status"])||"Backlog",priority:text(row,["prioridade"])||"Média",effort,progress:0,due:dueDate?new Date(dueDate+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"short"}).replace(".",""):"Sem prazo",dueDate,startDate:iso(text(row,["data de início","data de inicio","início","inicio"])),risk:"Médio",type:text(row,["categoria","tipo"])||"Evolutiva",version:text(row,["versão","versao"]),notes:text(row,["observações","observacoes"]),resourceIds:[]})});if(imported.length)setRows((current)=>[...imported,...current]);setImportFeedback(`${imported.length} demanda(s) importada(s)${ignored?` · ${ignored} linha(s) ignorada(s)`:""}.`)}catch{setImportFeedback("Não foi possível ler a planilha. Verifique o formato e os cabeçalhos.")}finally{setImporting(false);e.target.value=""}};
  const close = () => {
    setModal(false);
    setEditing(null);
    setErrors({});
    setForm(emptyDemandForm);
  };
  const openNew = () => {
    setEditing(null);
    setForm(emptyDemandForm);
    setModal(true);
  };
  const openEdit = (d: Demand) => {
    const month: Record<string, string> = {
      jan: "01",
      fev: "02",
      mar: "03",
      abr: "04",
      mai: "05",
      jun: "06",
      jul: "07",
      ago: "08",
      set: "09",
      out: "10",
      nov: "11",
      dez: "12",
    };
    const parts = d.due.toLowerCase().split(" ");
    setEditing(d);
    setForm({
      ...emptyDemandForm,
      name: d.name,
      product: d.product,
      priority: d.priority,
      status: d.status,
      type: d.type || "Evolutiva",
      version: d.version || "",
      description: d.notes || "",
      effort: String(d.effort),
      owner: d.owner,
      requester: d.owner,
      startDate: d.startDate || "",
      due: d.dueDate || (parts.length === 2 ? `2026-${month[parts[1]] || "09"}-${parts[0].padStart(2, "0")}` : ""),
      resourceIds: d.resourceIds || [],
    });
    setModal(true);
  };
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    const next: Record<string, string> = {};
    if (!form.name.trim()) next.name = "Informe o título da demanda.";
    if (!form.product) next.product = "Selecione um produto cadastrado.";
    if (!form.requester.trim()) next.requester = "Informe o solicitante.";
    if (!form.due) next.due = "Informe o prazo desejado.";
    if (Number(form.effort) <= 0) next.effort = "Informe um esforço válido.";
    setErrors(next);
    if (Object.keys(next).length) return;
    const date = new Date(form.due + "T12:00:00");
    const due = date
      .toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })
      .replace(".", "");
    const data: Demand = {
      id: editing?.id || `DEV-${160 + rows.length}`,
      name: form.name.trim(),
      product: form.product,
      owner: selectedProduct?.coordinator.split(" ")[0] || "Time",
      status: form.status,
      priority: form.priority,
      effort: Number(form.effort),
      progress: editing?.progress || 0,
      due,
      dueDate: form.due,
      risk: form.priority === "Crítica" ? "Alto" : editing?.risk || "Médio",
      type: form.type,
      version: form.version,
      notes: form.description,
      startDate: form.startDate,
      resourceIds: form.resourceIds,
    };
    setRows(
      editing
        ? rows.map((row) => (row.id === editing.id ? data : row))
        : [data, ...rows],
    );
    if(!editing && form.planningQuarter!=="Planejamento"){
      setRoadmap((current)=>[...current,{demandId:data.id,quarter:form.planningQuarter,collaborators:[],allocations:[]}]);
      setHistory((current)=>[{id:Date.now(),date:new Date().toLocaleString("pt-BR"),title:"Demanda planejada no cadastro",actor:"Vagner Moraes",detail:`${data.id} incluída em ${form.planningQuarter}`,tone:"good",product:data.product,quarter:form.planningQuarter},...current]);
    }
    close();
    setSaved(true);
    setTimeout(() => setSaved(false), 3500);
  };
  const confirmRemove = () => {
    if (!removing) return;
    setRows(rows.filter((row) => row.id !== removing.id));
    setRemoving(null);
  };
  return (
    <>
      <PageHead
        title="Gestão de demandas"
        desc="Priorize, acompanhe e entenda o impacto de cada entrega."
        action={
          <button className="primary" onClick={openNew}>
            <I.Plus /> Nova demanda
          </button>
        }
      />
      <div className="toolbar">
        <label className="search inner">
          <I.Search />
          <input
            placeholder="Buscar demandas..."
            value={q}
            onChange={(e) => setQ(e.target.value)}
          />
        </label>
        <label className="demand-filter"><I.Boxes/><select aria-label="Filtrar por produto" value={productFilter} onChange={(e)=>setProductFilter(e.target.value)}><option>Todos</option>{products.filter((product)=>product.active).map((product)=><option key={product.id} value={product.name}>{product.name}</option>)}</select></label>
        <label className="demand-filter"><I.CalendarRange/><select aria-label="Filtrar por quarter" value={quarterFilter} onChange={(e)=>setQuarterFilter(e.target.value)}><option>Todos</option><option>Planejamento</option>{quarters.map((quarter)=><option key={quarter}>{quarter}</option>)}</select></label>
        {(productFilter!=="Todos"||quarterFilter!=="Todos")&&<button onClick={()=>{setProductFilter("Todos");setQuarterFilter("Todos")}}><I.X/> Limpar</button>}
        <button onClick={()=>fileRef.current?.click()} disabled={importing}>
          {importing?<I.LoaderCircle className="spin"/>:<I.FileSpreadsheet/>} {importing?"Importando...":"Importar Excel"}
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={importExcel}/>
      </div>
      {importFeedback&&<div className="import-feedback"><I.CircleCheck/><span>{importFeedback}</span><button onClick={()=>setImportFeedback("")}><I.X/></button></div>}
      <Card>
        <div className="cardhead">
          <div>
            <h3>Todas as demandas</h3>
            <p>{filtered.length} resultados · atualizado agora</p>
          </div>
          <div className="seg">
            <button className="active">
              <I.List /> Lista
            </button>
            <button>
              <I.Columns3 /> Board
            </button>
          </div>
        </div>
        <DemandTableRows
          rows={filtered}
          onEdit={openEdit}
          onRemove={setRemoving}
        />
      </Card>
      {modal && (
        <div className="overlay demand-overlay" onMouseDown={close}>
          <form
            className="demand-modal"
            onMouseDown={(e) => e.stopPropagation()}
            onSubmit={submit}
          >
            <div className="modal-head">
              <div>
                <span className="modal-icon">
                  <I.ListPlus />
                </span>
                <div>
                  <h2>Nova demanda</h2>
                  <p>
                    Cadastre a demanda para avaliar seu impacto no planejamento.
                  </p>
                </div>
              </div>
              <button
                type="button"
                className="modal-close"
                onClick={close}
                aria-label="Fechar"
              >
                <I.X />
              </button>
            </div>
            <div className="modal-body">
              <div className="form-field wide">
                <label htmlFor="name">
                  Título <em>*</em>
                </label>
                <input
                  id="name"
                  name="name"
                  value={form.name}
                  onChange={change}
                  placeholder="Ex.: Integração com novo gateway"
                  autoFocus
                  className={errors.name ? "invalid" : ""}
                />
                {errors.name && (
                  <small className="field-error">{errors.name}</small>
                )}
              </div>
              <div className="form-field wide">
                <label htmlFor="description">Descrição</label>
                <textarea
                  id="description"
                  name="description"
                  value={form.description}
                  onChange={change}
                  placeholder="Descreva o contexto e o resultado esperado"
                  rows={3}
                />
              </div>
              <div className="form-field">
                <label>
                  Produto <em>*</em>
                </label>
                <select
                  name="product"
                  value={form.product}
                  onChange={change}
                  className={errors.product ? "invalid" : ""}
                >
                  <option value="">Selecione um produto...</option>
                  {products
                    .filter((product) => product.active)
                    .map((product) => (
                      <option key={product.id} value={product.name}>
                        {product.name}
                      </option>
                    ))}
                </select>
                {errors.product && (
                  <small className="field-error">{errors.product}</small>
                )}
              </div>
              <div className="form-field team-readonly"><label>Time responsável <I.Lock/></label><div><I.UsersRound/><span><b>{selectedProduct?.team || "Selecione um produto"}</b><small>{selectedProduct ? `Coordenador: ${selectedProduct.coordinator}` : "Preenchido automaticamente"}</small></span></div></div>
              <div className="form-field wide"><label>Recursos</label><div className="resource-selector">{!form.product?<p>Selecione um produto para visualizar os colaboradores da squad.</p>:availableResources.length===0?<p>Nenhum colaborador cadastrado para {selectedProduct?.team}.</p>:availableResources.map((person)=><label key={person.id} className={form.resourceIds.includes(person.id)?"selected":""}><input type="checkbox" checked={form.resourceIds.includes(person.id)} onChange={()=>setForm({...form,resourceIds:form.resourceIds.includes(person.id)?form.resourceIds.filter((id)=>id!==person.id):[...form.resourceIds,person.id]})}/><span className="avatar">{person.name.split(" ").map((x)=>x[0]).slice(0,2).join("")}</span><span><b>{person.name}</b><small>{person.role} · {person.total-person.used}h disponíveis</small></span></label>)}</div><small className="resource-help">{form.resourceIds.length} recurso(s) selecionado(s)</small></div>
              <div className="form-field">
                <label>
                  Categoria <em>*</em>
                </label>
                <select name="type" value={form.type} onChange={change}>
                  <option>Estratégica</option>
                  <option>Evolutiva</option>
                  <option>Bug</option>
                  <option>Incidente</option>
                  <option>Sustentação</option>
                  <option>Regulatória</option>
                  <option>Técnica</option>
                  <option>Emergencial</option>
                </select>
              </div>
              <div className="form-field">
                <label>Origem</label>
                <select name="origin" value={form.origin} onChange={change}>
                  <option>Produto</option>
                  <option>Cliente</option>
                  <option>Tecnologia</option>
                  <option>Regulatório</option>
                  <option>Operações</option>
                </select>
              </div>
              <div className="form-field"><label>Versão</label><input name="version" value={form.version} onChange={change} placeholder="Ex.: 5.2"/></div>
              <div className="form-field">
                <label htmlFor="requester">
                  Solicitante <em>*</em>
                </label>
                <input
                  id="requester"
                  name="requester"
                  value={form.requester}
                  onChange={change}
                  placeholder="Nome do solicitante"
                  className={errors.requester ? "invalid" : ""}
                />
                {errors.requester && (
                  <small className="field-error">{errors.requester}</small>
                )}
              </div>
              <div className="form-field">
                <label>Prioridade</label>
                <select name="priority" value={form.priority} onChange={change}>
                  <option>Baixa</option>
                  <option>Média</option>
                  <option>Alta</option>
                  <option>Crítica</option>
                </select>
              </div>
              <div className="form-field"><label>Status</label><select name="status" value={form.status} onChange={change}><option>Backlog</option><option>Planejada</option><option>Em análise</option><option>Desenvolvimento</option><option>Em testes</option><option>Homologação</option><option>Concluída</option><option>Cancelada</option><option>Bloqueada</option></select></div>
              {!editing&&<div className="form-field"><label>Planejamento inicial</label><select name="planningQuarter" value={form.planningQuarter} onChange={change}><option value="Planejamento">Deixar em planejamento</option>{quarters.map((quarter)=><option key={quarter}>{quarter}</option>)}</select><small className="resource-help">O quarter pertence ao planejamento, não à demanda.</small></div>}
              <div className="form-field">
                <label htmlFor="startDate">Data de início</label>
                <input id="startDate" type="date" name="startDate" value={form.startDate} onChange={change}/>
              </div>
              <div className="form-field">
                <label htmlFor="due">
                  Prazo desejado <em>*</em>
                </label>
                <input
                  id="due"
                  type="date"
                  name="due"
                  value={form.due}
                  onChange={change}
                  className={errors.due ? "invalid" : ""}
                />
                {errors.due && (
                  <small className="field-error">{errors.due}</small>
                )}
              </div>
              <div className="form-field">
                <label htmlFor="effort">Esforço estimado (horas)</label>
                <input
                  id="effort"
                  type="number"
                  min="1"
                  name="effort"
                  value={form.effort}
                  onChange={change}
                  className={errors.effort ? "invalid" : ""}
                />
                {errors.effort && (
                  <small className="field-error">{errors.effort}</small>
                )}
              </div>
              <div className="form-field wide">
                <label>Dependências</label>
                <input
                  name="dependencies"
                  value={form.dependencies}
                  onChange={change}
                  placeholder="Outros times, sistemas ou demandas"
                />
              </div>
              <div className="form-field">
                <label>Justificativa</label>
                <textarea
                  name="justification"
                  value={form.justification}
                  onChange={change}
                  rows={2}
                  placeholder="Por que esta demanda é necessária?"
                />
              </div>
              <div className="form-field">
                <label>Impacto esperado</label>
                <textarea
                  name="impact"
                  value={form.impact}
                  onChange={change}
                  rows={2}
                  placeholder="Benefícios e resultados esperados"
                />
              </div>
            </div>
            <div className="capacity-preview">
              <I.Gauge />
              <div>
                <b>Impacto preliminar na capacidade</b>
                <p>
                  A demanda consumirá {form.effort || 0}h. O impacto detalhado
                  será calculado após o cadastro.
                </p>
              </div>
            </div>
            <div className="modal-actions">
              <button type="button" className="ghost" onClick={close}>
                Cancelar
              </button>
              <button type="submit" className="primary">
                <I.Plus /> Cadastrar demanda
              </button>
            </div>
          </form>
        </div>
      )}
      {removing && (
        <div
          className="overlay confirm-overlay"
          onMouseDown={() => setRemoving(null)}
        >
          <div
            className="confirm-modal"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <span>
              <I.Trash2 />
            </span>
            <h3>Excluir demanda?</h3>
            <p>
              A demanda{" "}
              <b>
                {removing.id} — {removing.name}
              </b>{" "}
              será removida da lista. Esta ação não poderá ser desfeita.
            </p>
            <div>
              <button className="ghost" onClick={() => setRemoving(null)}>
                Cancelar
              </button>
              <button className="danger" onClick={confirmRemove}>
                <I.Trash2 /> Excluir demanda
              </button>
            </div>
          </div>
        </div>
      )}
      {saved && (
        <div className="toast-success">
          <I.CircleCheck />
          <div>
            <b>Demanda salva</b>
            <span>As informações foram atualizadas com sucesso.</span>
          </div>
          <button onClick={() => setSaved(false)}>
            <I.X />
          </button>
        </div>
      )}
    </>
  );
}
function DemandTableRows({
  rows,
  onEdit,
  onRemove,
}: {
  rows: Demand[];
  onEdit?: (d: Demand) => void;
  onRemove?: (d: Demand) => void;
}) {
  return (
    <div className="tablewrap">
      <table>
        <thead>
          <tr>
            <th>Demanda</th>
            <th>Produto</th>
            <th>Responsável</th>
            <th>Status</th>
            <th>Prioridade</th>
            <th>Progresso</th>
            <th>Prazo</th>
            <th>Risco</th>
            {onEdit && <th className="actions-head">Ações</th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((d) => (
            <tr key={d.id}>
              <td>
                <b>{d.name}</b>
                <small>{d.id}</small>
              </td>
              <td>{d.product}</td>
              <td>
                <span className="person">{d.owner[0]}</span>
                {d.owner}
              </td>
              <td>
                <Badge
                  tone={
                    d.status === "Bloqueada"
                      ? "bad"
                      : d.status === "Concluída"
                        ? "good"
                        : "blue"
                  }
                >
                  {d.status}
                </Badge>
              </td>
              <td>{d.priority}</td>
              <td>
                <div className="mini">
                  <i style={{ width: d.progress + "%" }} />
                </div>
                <small>{d.progress}%</small>
              </td>
              <td>{d.due}</td>
              <td>
                <Badge
                  tone={
                    d.risk === "Crítico"
                      ? "bad"
                      : d.risk === "Alto"
                        ? "warn"
                        : d.risk === "Baixo"
                          ? "good"
                          : "gray"
                  }
                >
                  {d.risk}
                </Badge>
              </td>
              {onEdit && (
                <td>
                  <div className="row-actions">
                    <button
                      title="Editar demanda"
                      aria-label={`Editar ${d.id}`}
                      onClick={() => onEdit(d)}
                    >
                      <I.Pencil />
                    </button>
                    <button
                      className="delete"
                      title="Excluir demanda"
                      aria-label={`Excluir ${d.id}`}
                      onClick={() => onRemove?.(d)}
                    >
                      <I.Trash2 />
                    </button>
                  </div>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const emptyProduct = { name: "", code: "", description: "", team: "", coordinator: "", active: true };
function Products() {
  const { products, setProducts } = useProducts();
  const { rows } = useDemands();
  const { team } = useTeam();
  const [modal, setModal] = useState(false);
  const [editing, setEditing] = useState<Product | null>(null);
  const [form, setForm] = useState(emptyProduct);
  const [error, setError] = useState("");
  const openNew = () => {
    setEditing(null);
    setForm(emptyProduct);
    setError("");
    setModal(true);
  };
  const openEdit = (product: Product) => {
    setEditing(product);
    setForm({
      name: product.name,
      code: product.code,
      description: product.description,
      team: product.team,
      coordinator: product.coordinator,
      active: product.active,
    });
    setError("");
    setModal(true);
  };
  const close = () => {
    setModal(false);
    setEditing(null);
    setError("");
  };
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.name.trim() || !form.code.trim() || !form.team.trim() || !form.coordinator.trim()) {
      setError("Nome, código, time e coordenador são obrigatórios.");
      return;
    }
    if (
      products.some(
        (p) =>
          p.id !== editing?.id &&
          (p.name.toLowerCase() === form.name.trim().toLowerCase() ||
            p.code.toLowerCase() === form.code.trim().toLowerCase()),
      )
    ) {
      setError("Já existe um produto com esse nome ou código.");
      return;
    }
    const data: Product = {
      id: editing?.id || Date.now(),
      name: form.name.trim(),
      code: form.code.trim().toUpperCase(),
      description: form.description.trim(),
      team: form.team.trim(),
      coordinator: form.coordinator.trim(),
      active: form.active,
    };
    setProducts(
      editing
        ? products.map((p) => (p.id === editing.id ? data : p))
        : [...products, data],
    );
    close();
  };
  const remove = (product: Product) => {
    const usage =
      rows.filter((d) => d.product === product.name).length +
      team.filter((m) => m.product === product.name).length;
    if (usage) {
      setError(
        `O produto ${product.name} não pode ser excluído porque possui ${usage} vínculo(s).`,
      );
      return;
    }
    setProducts(products.filter((p) => p.id !== product.id));
  };
  return (
    <>
      <PageHead
        title="Produtos"
        desc="Cadastre os produtos utilizados por demandas e colaboradores."
        action={
          <button className="primary" onClick={openNew}>
            <I.Plus /> Novo produto
          </button>
        }
      />
      <div className="product-summary">
        <I.Database />
        <span>
          <b>Cadastro mestre</b> Demandas e colaboradores só podem utilizar
          produtos ativos desta lista.
        </span>
        <Badge tone="good">
          {products.filter((p) => p.active).length} ativos
        </Badge>
      </div>
      <div className="product-grid">
        {products.map((product) => {
          const demandsCount = rows.filter(
            (d) => d.product === product.name,
          ).length;
          const staffCount = team.filter(
            (m) => m.product === product.name,
          ).length;
          return (
            <Card className="product-card" key={product.id}>
              <div className="product-code">{product.code}</div>
              <div className="product-info">
                <div>
                  <h3>{product.name}</h3>
                  <Badge tone={product.active ? "good" : "gray"}>
                    {product.active ? "Ativo" : "Inativo"}
                  </Badge>
                </div>
                <p>
                  <I.UsersRound /> {product.team}
                </p>
                <p><I.UserRound /> Coordenador: {product.coordinator}</p>
                <div>
                  <span>
                    <b>{demandsCount}</b> demandas
                  </span>
                  <span>
                    <b>{staffCount}</b> colaboradores
                  </span>
                </div>
              </div>
              <div className="row-actions">
                <button
                  title="Editar produto"
                  onClick={() => openEdit(product)}
                >
                  <I.Pencil />
                </button>
                <button
                  className="delete"
                  title="Excluir produto"
                  onClick={() => remove(product)}
                >
                  <I.Trash2 />
                </button>
              </div>
            </Card>
          );
        })}
      </div>
      {error && !modal && (
        <div className="toast-success product-error">
          <I.CircleAlert />
          <div>
            <b>Produto em uso</b>
            <span>{error}</span>
          </div>
          <button onClick={() => setError("")}>
            <I.X />
          </button>
        </div>
      )}
      {modal && (
        <div className="overlay confirm-overlay" onMouseDown={close}>
          <form
            className="staff-modal"
            onMouseDown={(e) => e.stopPropagation()}
            onSubmit={submit}
          >
            <div className="modal-head">
              <div>
                <span className="modal-icon">
                  <I.Boxes />
                </span>
                <div>
                  <h2>{editing ? "Editar produto" : "Novo produto"}</h2>
                  <p>Este produto ficará disponível nos demais cadastros.</p>
                </div>
              </div>
              <button type="button" className="modal-close" onClick={close}>
                <I.X />
              </button>
            </div>
            <div className="staff-form">
              <div className="form-field">
                <label>
                  Nome <em>*</em>
                </label>
                <input
                  autoFocus
                  value={form.name}
                  onChange={(e) => setForm({ ...form, name: e.target.value })}
                  placeholder="Ex.: Marketplace"
                />
              </div>
              <div className="form-field">
                <label>
                  Código <em>*</em>
                </label>
                <input
                  value={form.code}
                  maxLength={6}
                  onChange={(e) => setForm({ ...form, code: e.target.value })}
                  placeholder="Ex.: MKT"
                />
              </div>
              <div className="form-field wide"><label>Descrição</label><textarea rows={2} value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} placeholder="Contexto e objetivo do produto"/></div>
              <div className="form-field"><label>Time responsável <em>*</em></label><input value={form.team} onChange={(e) => setForm({ ...form, team: e.target.value })} placeholder="Ex.: Squad Marketplace"/></div>
              <div className="form-field"><label>Coordenador <em>*</em></label><input value={form.coordinator} onChange={(e) => setForm({ ...form, coordinator: e.target.value })} placeholder="Nome do coordenador"/></div>
              <label className="product-active">
                <input
                  type="checkbox"
                  checked={form.active}
                  onChange={(e) =>
                    setForm({ ...form, active: e.target.checked })
                  }
                />
                <span>
                  <b>Produto ativo</b>
                  <small>Disponível para novas demandas e colaboradores</small>
                </span>
              </label>
              {error && (
                <div className="staff-error">
                  <I.CircleAlert />
                  {error}
                </div>
              )}
            </div>
            <div className="modal-actions">
              <button type="button" className="ghost" onClick={close}>
                Cancelar
              </button>
              <button className="primary" type="submit">
                <I.Save /> Salvar produto
              </button>
            </div>
          </form>
        </div>
      )}
    </>
  );
}

type Staff = {
  id: number;
  name: string;
  role: string;
  product: string;
  total: number;
  used: number;
};
const initialStaff: Staff[] = members.map(([name, role, total, used], i) => ({
  id: i + 1,
  name: String(name),
  role: String(role),
  product: [
    "Commerce",
    "Payments",
    "Customer",
    "Payments",
    "Platform",
    "Analytics",
  ][i],
  total: Number(total),
  used: Number(used),
}));
const TeamContext = createContext<{
  team: Staff[];
  setTeam: React.Dispatch<React.SetStateAction<Staff[]>>;
} | null>(null);
function TeamProvider({ children }: { children: React.ReactNode }) {
  const [team, setTeam] = useState<Staff[]>(initialStaff);
  return (
    <TeamContext.Provider value={{ team, setTeam }}>
      {children}
    </TeamContext.Provider>
  );
}
function useTeam() {
  const value = useContext(TeamContext);
  if (!value) throw new Error("TeamProvider ausente");
  return value;
}
const emptyStaff = {
  name: "",
  role: "Backend",
  product: "",
  total: "130",
  used: "0",
};
function Capacity() {
  const { team, setTeam } = useTeam();
  const { products } = useProducts();
  const { rows, roadmap } = useDemands();
  const [productFilter,setProductFilter]=useState("Todos");
  const [allocationFor,setAllocationFor]=useState<Staff|null>(null);
  const [form, setForm] = useState(emptyStaff);
  const [editing, setEditing] = useState<Staff | null>(null);
  const [removing, setRemoving] = useState<Staff | null>(null);
  const [modal, setModal] = useState(false);
  const [error, setError] = useState("");
  const filteredTeam=productFilter==="Todos"?team:team.filter((person)=>person.product===productFilter);
  const total = filteredTeam.reduce((sum, m) => sum + m.total, 0);
  const used = filteredTeam.reduce((sum, m) => sum + m.used, 0);
  const free = total - used;
  const overloaded = filteredTeam.filter((m) => m.used / m.total > 0.9).length;
  const allocatedDemands=allocationFor?roadmap.flatMap((plan)=>{const allocation=plan.allocations.find((item)=>item.staffId===allocationFor.id);const demand=rows.find((item)=>item.id===plan.demandId);return allocation&&demand?[{demand,plan,hours:allocation.hours}]:[]}):[];
  const openNew = () => {
    setEditing(null);
    setForm(emptyStaff);
    setError("");
    setModal(true);
  };
  const openEdit = (m: Staff) => {
    setEditing(m);
    setForm({
      name: m.name,
      role: m.role,
      product: m.product,
      total: String(m.total),
      used: String(m.used),
    });
    setError("");
    setModal(true);
  };
  const close = () => {
    setModal(false);
    setEditing(null);
    setError("");
  };
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    const cap = Number(form.total),
      allocation = Number(form.used);
    if (!form.name.trim()) {
      setError("Informe o nome do colaborador.");
      return;
    }
    if (!form.product) {
      setError("Selecione o produto do colaborador.");
      return;
    }
    if (cap <= 0 || allocation < 0) {
      setError("Informe valores válidos para capacidade e alocação.");
      return;
    }
    const data: Staff = {
      id: editing?.id || Date.now(),
      name: form.name.trim(),
      role: form.role,
      product: form.product,
      total: cap,
      used: allocation,
    };
    setTeam(
      editing
        ? team.map((m) => (m.id === editing.id ? data : m))
        : [...team, data],
    );
    close();
  };
  const confirmRemove = () => {
    if (removing) setTeam(team.filter((m) => m.id !== removing.id));
    setRemoving(null);
  };
  return (
    <>
      <PageHead
        title="Capacidade do time"
        desc="Visualize alocação, disponibilidade e pontos de sobrecarga."
        action={
          <button className="primary" onClick={openNew}>
            <I.UserPlus /> Novo colaborador
          </button>
        }
      />
      <div className="capacity-filter"><div><I.Boxes/><span><b>Capacidade por produto</b><small>Filtre os indicadores e colaboradores</small></span></div><select value={productFilter} onChange={(e)=>setProductFilter(e.target.value)}><option>Todos</option>{products.filter((product)=>product.active).map((product)=><option key={product.id} value={product.name}>{product.name}</option>)}</select></div>
      <div className="kpis compact">
        <Card>
          <span>Capacidade total</span>
          <strong>{total}h</strong>
          <small>{filteredTeam.length} colaboradores</small>
        </Card>
        <Card>
          <span>Alocada</span>
          <strong>{used}h</strong>
          <small className="warn">
            {total ? Math.round((used / total) * 100) : 0}% utilizada
          </small>
        </Card>
        <Card>
          <span>Disponível</span>
          <strong className={free < 0 ? "bad" : ""}>{free}h</strong>
          <small className={free < 0 ? "bad" : "good"}>
            {free < 0 ? "Capacidade excedida" : "Reserva disponível"}
          </small>
        </Card>
        <Card>
          <span>Sobrecarregados</span>
          <strong>{overloaded}</strong>
          <small className={overloaded ? "bad" : "good"}>
            {overloaded ? "Requer atenção" : "Time equilibrado"}
          </small>
        </Card>
      </div>
      <Card>
        <div className="cardhead">
          <div>
            <h3>Alocação por colaborador</h3>
            <p>Capacidade planejada para setembro</p>
          </div>
          <Badge tone="gray">{filteredTeam.length} pessoas</Badge>
        </div>
        <div className="people">
          {filteredTeam.map((m) => {
            const pct = Math.round((m.used / m.total) * 100);
            return (
              <div className="member member-managed" key={m.id}>
                <span className="avatar">
                  {m.name
                    .split(" ")
                    .map((x) => x[0])
                    .slice(0, 2)
                    .join("")}
                </span>
                <div>
                  <b>{m.name}</b>
                  <small>
                    {m.role} · {m.product}
                  </small>
                </div>
                <div className="usage">
                  <div className="usage-bar" role="button" tabIndex={0} title="Ver demandas alocadas" onClick={()=>setAllocationFor(m)} onKeyDown={(e)=>{if(e.key==="Enter")setAllocationFor(m)}}>
                    <i
                      className={pct > 90 ? "hot" : ""}
                      style={{ width: Math.min(pct, 100) + "%" }}
                    />
                  </div>
                  <small>
                    {m.used}h de {m.total}h
                  </small>
                </div>
                <b className={pct > 90 ? "bad" : ""}>{pct}%</b>
                <span className={m.total - m.used < 0 ? "bad" : ""}>
                  {m.total - m.used}h livres
                </span>
                <div className="row-actions">
                  <button
                    title="Editar colaborador"
                    onClick={() => openEdit(m)}
                  >
                    <I.Pencil />
                  </button>
                  <button
                    className="delete"
                    title="Excluir colaborador"
                    onClick={() => setRemoving(m)}
                  >
                    <I.Trash2 />
                  </button>
                </div>
              </div>
            );
          })}
          {filteredTeam.length === 0 && (
            <div className="team-empty">
              <I.Users />
              <b>Nenhum colaborador cadastrado</b>
              <button className="primary" onClick={openNew}>
                Adicionar colaborador
              </button>
            </div>
          )}
        </div>
      </Card>
      {allocationFor&&<div className="overlay confirm-overlay" onMouseDown={()=>setAllocationFor(null)}><div className="capacity-demand-modal" onMouseDown={(e)=>e.stopPropagation()}><div className="modal-head"><div><span className="modal-icon"><I.ListTodo/></span><div><h2>Demandas alocadas</h2><p>{allocationFor.name} · {allocationFor.product}</p></div></div><button className="modal-close" onClick={()=>setAllocationFor(null)}><I.X/></button></div><div className="capacity-demand-list">{allocatedDemands.length===0?<div className="picker-empty"><I.CalendarX/><b>Nenhuma demanda alocada</b><span>Não existem alocações no Roadmap para este colaborador.</span></div>:allocatedDemands.map(({demand,plan,hours})=><div key={demand.id}><span className="demand-dot"/><div><b>{demand.name}</b><small>{demand.id} · {demand.product} · {plan.quarter}</small></div><Badge tone="blue">{hours}h</Badge></div>)}</div><div className="allocation-total"><span>Total nas demandas</span><b>{allocatedDemands.reduce((sum,item)=>sum+item.hours,0)}h</b></div><div className="picker-foot"><button className="primary" onClick={()=>setAllocationFor(null)}>Fechar</button></div></div></div>}
      {modal && (
        <div className="overlay confirm-overlay" onMouseDown={close}>
          <form
            className="staff-modal"
            onMouseDown={(e) => e.stopPropagation()}
            onSubmit={submit}
          >
            <div className="modal-head">
              <div>
                <span className="modal-icon">
                  {editing ? <I.UserRoundPen /> : <I.UserPlus />}
                </span>
                <div>
                  <h2>{editing ? "Editar colaborador" : "Novo colaborador"}</h2>
                  <p>Defina o perfil e a capacidade disponível no período.</p>
                </div>
              </div>
              <button type="button" className="modal-close" onClick={close}>
                <I.X />
              </button>
            </div>
            <div className="staff-form">
              <div className="form-field wide">
                <label>
                  Nome completo <em>*</em>
                </label>
                <input
                  autoFocus
                  value={form.name}
                  onChange={(e) => setForm({ ...form, name: e.target.value })}
                  placeholder="Ex.: Ana Oliveira"
                  className={error && !form.name ? "invalid" : ""}
                />
              </div>
              <div className="form-field wide">
                <label>Função</label>
                <select
                  value={form.role}
                  onChange={(e) => setForm({ ...form, role: e.target.value })}
                >
                  <option>Tech Lead</option>
                  <option>Backend</option>
                  <option>Frontend</option>
                  <option>Full Stack</option>
                  <option>UX Engineer</option>
                  <option>QA Engineer</option>
                  <option>Product Manager</option>
                </select>
              </div>
              <div className="form-field wide">
                <label>
                  Produto <em>*</em>
                </label>
                <select
                  value={form.product}
                  onChange={(e) =>
                    setForm({ ...form, product: e.target.value })
                  }
                >
                  <option value="">Selecione um produto...</option>
                  {products
                    .filter((product) => product.active)
                    .map((product) => (
                      <option key={product.id} value={product.name}>
                        {product.name}
                      </option>
                    ))}
                </select>
              </div>
              <div className="form-field">
                <label>Capacidade (horas)</label>
                <input
                  type="number"
                  min="1"
                  value={form.total}
                  onChange={(e) => setForm({ ...form, total: e.target.value })}
                />
              </div>
              <div className="form-field">
                <label>Horas alocadas</label>
                <input
                  type="number"
                  min="0"
                  value={form.used}
                  onChange={(e) => setForm({ ...form, used: e.target.value })}
                />
              </div>
              {error && (
                <div className="staff-error">
                  <I.CircleAlert />
                  {error}
                </div>
              )}
            </div>
            <div className="modal-actions">
              <button type="button" className="ghost" onClick={close}>
                Cancelar
              </button>
              <button className="primary" type="submit">
                {editing ? <I.Save /> : <I.UserPlus />}
                {editing ? "Salvar alterações" : "Adicionar colaborador"}
              </button>
            </div>
          </form>
        </div>
      )}
      {removing && (
        <div
          className="overlay confirm-overlay"
          onMouseDown={() => setRemoving(null)}
        >
          <div
            className="confirm-modal"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <span>
              <I.UserRoundX />
            </span>
            <h3>Excluir colaborador?</h3>
            <p>
              <b>{removing.name}</b> será removido do planejamento de
              capacidade. Esta ação não poderá ser desfeita.
            </p>
            <div>
              <button className="ghost" onClick={() => setRemoving(null)}>
                Cancelar
              </button>
              <button className="danger" onClick={confirmRemove}>
                <I.Trash2 /> Excluir colaborador
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function Planning() {
  const [hours, setHours] = useState(60);
  const occupancy = Math.round(((852 + hours) / 1040) * 100);
  return (
    <>
      <PageHead
        title="Simulação de cenários"
        desc="Teste decisões antes de comprometer o planejamento real."
      />
      <div className="scenario">
        <Card>
          <div className="cardhead">
            <div>
              <h3>Configurar mudança</h3>
              <p>Adicione uma demanda hipotética</p>
            </div>
          </div>
          <label>
            Título
            <input defaultValue="Nova demanda estratégica" />
          </label>
          <div className="form2">
            <label>
              Tipo
              <select>
                <option>Estratégica</option>
                <option>Emergencial</option>
              </select>
            </label>
            <label>
              Prioridade
              <select>
                <option>Alta</option>
                <option>Crítica</option>
              </select>
            </label>
          </div>
          <label>
            Esforço estimado <b>{hours}h</b>
            <input
              type="range"
              min="10"
              max="180"
              value={hours}
              onChange={(e) => setHours(+e.target.value)}
            />
          </label>
          <button className="primary full">
            <I.Play /> Recalcular cenário
          </button>
        </Card>
        <Card className="impact">
          <div className="cardhead">
            <div>
              <h3>Impacto projetado</h3>
              <p>Comparação com o cenário atual</p>
            </div>
            <Badge tone={occupancy > 90 ? "bad" : "warn"}>
              {occupancy > 90 ? "Alto impacto" : "Atenção"}
            </Badge>
          </div>
          <div className="compare">
            <div>
              <span>CENÁRIO ATUAL</span>
              <strong>82%</strong>
              <small>852h comprometidas</small>
            </div>
            <I.ArrowRight />
            <div>
              <span>CENÁRIO SIMULADO</span>
              <strong className={occupancy > 90 ? "bad" : ""}>
                {occupancy}%
              </strong>
              <small>{852 + hours}h comprometidas</small>
            </div>
          </div>
          <div className="alert bad">
            <I.TriangleAlert />
            <div>
              <b>Capacidade próxima do limite</b>
              <p>A mudança pode impactar 2 entregas do roadmap.</p>
            </div>
          </div>
          <div className="impactrow">
            <span>
              <i className="dot bad" />
              <b>Portal de autoatendimento</b>
            </span>
            <b>+5 dias</b>
          </div>
          <div className="impactrow">
            <span>
              <i className="dot warn" />
              <b>Novo checkout omnichannel</b>
            </span>
            <b>+2 dias</b>
          </div>
          <button className="primary full">Aplicar ao planejamento</button>
        </Card>
      </div>
    </>
  );
}

function Risks() {
  return (
    <>
      <PageHead
        title="Mapa de riscos"
        desc="Antecipe ameaças e acompanhe os planos de mitigação."
        action={
          <button className="primary">
            <I.Plus /> Novo risco
          </button>
        }
      />
      <div className="riskgrid">
        <Card>
          <div className="cardhead">
            <div>
              <h3>Matriz de probabilidade × impacto</h3>
              <p>Clique em uma célula para filtrar</p>
            </div>
          </div>
          <div className="matrix">
            <span />
            <b>Baixo</b>
            <b>Médio</b>
            <b>Alto</b>
            <b>Crítico</b>
            {["Alta", "Média", "Baixa"].map((r, ri) => (
              <React.Fragment key={r}>
                <b>{r}</b>
                {[0, 1, 2, 3].map((_, ci) => (
                  <div className={"cell c" + Math.min(3, ci + (2 - ri))}>
                    {ri === 0 && ci === 2 ? (
                      <i>2</i>
                    ) : ri === 1 && ci === 3 ? (
                      <i>1</i>
                    ) : (
                      ""
                    )}
                  </div>
                ))}
              </React.Fragment>
            ))}
          </div>
        </Card>
        <Card>
          <div className="cardhead">
            <div>
              <h3>Riscos prioritários</h3>
              <p>Ordenados por exposição</p>
            </div>
          </div>
          {[
            "Dependência da API antifraude",
            "Capacidade de backend em setembro",
            "Prazo regulatório inegociável",
          ].map((x, i) => (
            <div className="riskitem">
              <span className={"risknum r" + i}>{i + 1}</span>
              <div>
                <b>{x}</b>
                <small>
                  {
                    [
                      "DEV-148 · Rafael Lima",
                      "Time Backend · Marina Costa",
                      "DEV-154 · Lucas Rocha",
                    ][i]
                  }
                </small>
              </div>
              <Badge tone={i === 2 ? "bad" : "warn"}>
                {i === 2 ? "Crítico" : "Alto"}
              </Badge>
            </div>
          ))}
        </Card>
      </div>
    </>
  );
}

function LegacyChanges() {
  return (
    <>
      <PageHead
        title="Mudanças do roadmap"
        desc="Rastreabilidade completa das decisões que alteraram o trimestre."
        action={
          <button className="primary">
            <I.Plus /> Registrar mudança
          </button>
        }
      />
      <Card>
        {[
          [
            "Hoje, 14:32",
            "DEV-154 adicionada ao trimestre",
            "Vagner Moraes",
            "+60h · impacto alto",
          ],
          [
            "Ontem, 16:10",
            "Prioridade do checkout alterada",
            "Marina Costa",
            "Alta → Crítica",
          ],
          [
            "29 Ago, 10:24",
            "Esforço do Portal revisado",
            "Camila Souza",
            "120h → 140h",
          ],
          [
            "27 Ago, 09:05",
            "Otimização de busca concluída",
            "Bruno Alves",
            "Entrega antecipada",
          ],
        ].map((x, i) => (
          <div className="change">
            <div className={"changeicon i" + i}>
              {i === 0 ? (
                <I.Plus />
              ) : i === 1 ? (
                <I.ArrowUp />
              ) : i === 2 ? (
                <I.Clock />
              ) : (
                <I.Check />
              )}
            </div>
            <div>
              <small>{x[0]}</small>
              <b>{x[1]}</b>
              <span>por {x[2]}</span>
            </div>
            <Badge tone={i === 0 ? "bad" : i === 3 ? "good" : "gray"}>
              {x[3]}
            </Badge>
          </div>
        ))}
      </Card>
    </>
  );
}
function Changes(){const {history}=useDemands();const {products}=useProducts();const [product,setProduct]=useState("Todos");const [quarter,setQuarter]=useState("Todos");const filtered=history.filter((item)=>(product==="Todos"||item.product===product)&&(quarter==="Todos"||item.quarter===quarter));return <><PageHead title="Mudanças do roadmap" desc="Histórico automático de inclusões, movimentações, remoções e alocações."/><div className="changes-filters"><div><I.Filter/><span><b>Filtrar histórico</b><small>{filtered.length} de {history.length} alterações</small></span></div><label>Produto<select value={product} onChange={(e)=>setProduct(e.target.value)}><option>Todos</option>{products.map((item)=><option key={item.id} value={item.name}>{item.name}</option>)}</select></label><label>Quarter<select value={quarter} onChange={(e)=>setQuarter(e.target.value)}><option>Todos</option>{quarters.map((item)=><option key={item}>{item}</option>)}</select></label>{(product!=="Todos"||quarter!=="Todos")&&<button onClick={()=>{setProduct("Todos");setQuarter("Todos")}}><I.X/> Limpar</button>}</div><Card>{filtered.length===0?<div className="changes-empty"><I.History/><b>Nenhuma alteração encontrada</b><span>{history.length?"Altere ou limpe os filtros para visualizar outros registros.":"As ações realizadas no Roadmap aparecerão automaticamente aqui."}</span></div>:filtered.map((item)=><div className="change" key={item.id}><div className={`changeicon log-${item.tone}`}>{item.tone==="good"?<I.Plus/>:item.tone==="bad"?<I.Trash2/>:item.tone==="warn"?<I.ArrowRightLeft/>:<I.Users/>}</div><div><small>{item.date}</small><b>{item.title}</b><span>por {item.actor}{item.product?` · ${item.product}`:""}{item.quarter?` · ${item.quarter}`:""}</span></div><Badge tone={item.tone}>{item.detail}</Badge></div>)}</Card></>}
function Reports() {
  return (
    <>
      <PageHead
        title="Relatórios"
        desc="Transforme dados do trimestre em decisões compartilháveis."
        action={
          <button className="primary">
            <I.Download /> Exportar relatório
          </button>
        }
      />
      <div className="reportgrid">
        {[
          [
            "Evolução do roadmap",
            "Progresso planejado versus realizado",
            I.TrendingUp,
          ],
          [
            "Capacidade e utilização",
            "Alocação por time, pessoa e período",
            I.Gauge,
          ],
          [
            "Mudanças do trimestre",
            "Inclusões, remoções e impacto acumulado",
            I.GitCompareArrows,
          ],
          [
            "Riscos e previsibilidade",
            "Exposição e tendência das entregas",
            I.ShieldCheck,
          ],
        ].map(([a, b, Icon]: any) => (
          <Card className="report">
            <span>
              <Icon />
            </span>
            <div>
              <h3>{a}</h3>
              <p>{b}</p>
            </div>
            <button className="icon">
              <I.ArrowUpRight />
            </button>
          </Card>
        ))}
      </div>
    </>
  );
}
function Settings() {
  return (
    <>
      <PageHead
        title="Configurações"
        desc="Personalize o workspace, times e regras de planejamento."
      />
      <Card>
        <div className="settings">
          <div>
            <h3>Preferências do planejamento</h3>
            <p>Defina limites e comportamentos padrão.</p>
          </div>
          <label>
            Limite de ocupação recomendado{" "}
            <select>
              <option>85%</option>
              <option>90%</option>
            </select>
          </label>
          <label>
            Reserva técnica mínima{" "}
            <select>
              <option>10%</option>
              <option>15%</option>
            </select>
          </label>
          <label className="switchrow">
            <span>
              <b>Alertas de capacidade</b>
              <small>Notificar quando um time exceder o limite</small>
            </span>
            <input type="checkbox" defaultChecked />
          </label>
          <button className="primary">Salvar alterações</button>
        </div>
      </Card>
    </>
  );
}

type AccessRole="Administrador"|"Editor"|"Visualização";
function UsersAdmin(){const [users,setUsers]=useState([{id:1,name:"Vagner Moraes",email:"vagner@atlas.com",role:"Administrador" as AccessRole},{id:2,name:"Marina Costa",email:"marina@atlas.com",role:"Editor" as AccessRole},{id:3,name:"Bianca Melo",email:"bianca@atlas.com",role:"Visualização" as AccessRole}]);return <><PageHead title="Administração de usuários" desc="Gerencie os níveis de acesso ao portal." action={<button className="primary"><I.UserPlus/> Novo usuário</button>}/><Card><div className="tablewrap"><table><thead><tr><th>Usuário</th><th>E-mail</th><th>Perfil</th><th>Permissões</th></tr></thead><tbody>{users.map((user)=><tr key={user.id}><td><span className="person">{user.name[0]}</span><b>{user.name}</b></td><td>{user.email}</td><td><select className="quarter-inline" value={user.role} onChange={(e)=>setUsers(users.map((item)=>item.id===user.id?{...item,role:e.target.value as AccessRole}:item))}><option>Administrador</option><option>Editor</option><option>Visualização</option></select></td><td>{user.role==="Administrador"?"Acesso total":user.role==="Editor"?"Demandas, produtos, capacidade e roadmap":"Somente demandas e roadmap"}</td></tr>)}</tbody></table></div></Card></>}

function App() {
  const [page, setPage] = useState<Page>("Visão geral");
  const [dark, setDark] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [cmd, setCmd] = useState(false);
  const [role,setRole]=useState<AccessRole>("Administrador");
  const allowedNav=nav.filter(([name])=>role==="Administrador"?true:role==="Editor"?["Visão geral","Demandas","Roadmap","Produtos","Capacidade"].includes(name):["Demandas","Roadmap"].includes(name));
  useEffect(()=>{if(!allowedNav.some(([name])=>name===page))setPage(role==="Visualização"?"Demandas":"Visão geral")},[role]);
  useEffect(() => {
    const f = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "k") {
        e.preventDefault();
        setCmd(true);
      }
    };
    addEventListener("keydown", f);
    return () => removeEventListener("keydown", f);
  }, []);
  return (
    <div className={`${dark ? "app dark" : "app"} ${role === "Visualização" ? "role-viewer" : role === "Editor" ? "role-editor" : "role-admin"}`}>
      <aside className={collapsed ? "collapsed" : ""}>
        <div className="brand">
          <span>
            <I.Waypoints />
          </span>
          <b>
            Road<span>map</span>
          </b>
          <button onClick={() => setCollapsed(!collapsed)}>
            <I.PanelLeftClose />
          </button>
        </div>
        <nav>
          {allowedNav.map(([n, Icon]) => (
            <button
              title={n}
              className={page === n ? "active" : ""}
              onClick={() => setPage(n)}
            >
              <Icon />
              <span>{n}</span>
              {n === "Riscos" && <em>3</em>}
            </button>
          ))}
        </nav>
        <div className="sidefoot">
          <div className="workspace">
            <span>AX</span>
            <div>
              <b>Atlas Digital</b>
              <small>Workspace enterprise</small>
            </div>
            <I.ChevronsUpDown />
          </div>
        </div>
      </aside>
      <main>
        <header>
          <button className="mobile" onClick={() => setCollapsed(!collapsed)}>
            <I.Menu />
          </button>
          <button className="topsearch" onClick={() => setCmd(true)}>
            <I.Search />
            <span>Buscar demandas, projetos, pessoas...</span>
            <kbd>Ctrl K</kbd>
          </button>
          <div className="topactions">
            <button onClick={() => setDark(!dark)}>
              {dark ? <I.Sun /> : <I.Moon />}
            </button>
            <button className="bell">
              <I.Bell />
              <i />
            </button>
            <div className="user">
              <span>VM</span>
              <div>
                <b>Vagner Moraes</b>
                <select className="role-switch" value={role} onChange={(e)=>setRole(e.target.value as AccessRole)}><option>Administrador</option><option>Editor</option><option>Visualização</option></select>
              </div>
              <I.ChevronDown />
            </div>
          </div>
        </header>
        <div className="content">
          {page === "Visão geral" ? (
            <Dashboard go={setPage} />
          ) : page === "Roadmap" ? (
            <Roadmap />
          ) : page === "Demandas" ? (
            <Demands />
          ) : page === "Produtos" ? (
            <Products />
          ) : page === "Capacidade" ? (
            <Capacity />
          ) : page === "Planejamento" ? (
            <Planning />
          ) : page === "Riscos" ? (
            <Risks />
          ) : page === "Mudanças" ? (
            <Changes />
          ) : page === "Relatórios" ? (
            <Reports />
          ) : page === "Usuários" ? (
            <UsersAdmin />
          ) : (
            <Settings />
          )}
        </div>
      </main>
      {cmd && (
        <div className="overlay" onMouseDown={() => setCmd(false)}>
          <div className="command" onMouseDown={(e) => e.stopPropagation()}>
            <label>
              <I.Search />
              <input autoFocus placeholder="O que você procura?" />
            </label>
            <small>NAVEGAÇÃO</small>
            {nav.slice(0, 6).map(([n, Icon]) => (
              <button
                onClick={() => {
                  setPage(n);
                  setCmd(false);
                }}
              >
                <Icon />
                {n}
                <I.ArrowRight />
              </button>
            ))}
            <footer>
              <span>↑↓ para navegar</span>
              <span>ESC para fechar</span>
            </footer>
          </div>
        </div>
      )}
    </div>
  );
}
createRoot(document.getElementById("root")!).render(
  <ProductProvider>
    <TeamProvider>
      <DemandProvider>
        <App />
      </DemandProvider>
    </TeamProvider>
  </ProductProvider>,
);
