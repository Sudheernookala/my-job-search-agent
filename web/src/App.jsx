import { useEffect, useMemo, useState } from "react";

const EXCEL_URL =
  "https://github.com/Sudheernookala/my-job-search-agent/raw/main/data/job_history.xlsx";

const QUICK_RANGES = [
  { label: "Today", days: 1 },
  { label: "3 days", days: 3 },
  { label: "7 days", days: 7 },
  { label: "14 days", days: 14 },
  { label: "All (30 days)", days: null },
];

function isoDaysAgo(n) {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() - n); // dates in data are UTC
  return d.toISOString().slice(0, 10);
}

function formatDate(iso) {
  if (!iso) return "";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" });
}

// Lowercase and strip accents so "munchen" matches "München"
function norm(v) {
  return String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function isLink(v) {
  return typeof v === "string" && /^https?:\/\//i.test(v);
}

export default function App() {
  const [data, setData] = useState({ jobs: [], generated_at: null });
  const [status, setStatus] = useState("loading");

  const [company, setCompany] = useState("");
  const [location, setLocation] = useState("");
  const [source, setSource] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [activeRange, setActiveRange] = useState(null);
  const [sort, setSort] = useState({ key: "date_found", dir: "desc" });

  useEffect(() => {
    fetch(`./jobs.json?t=${Date.now()}`)
      .then((r) => {
        if (!r.ok) throw new Error(r.statusText);
        return r.json();
      })
      .then((d) => {
        setData(d);
        setStatus("ready");
      })
      .catch(() => setStatus("error"));
  }, []);

  const jobs = data.jobs || [];
  const latestDate = jobs.reduce((max, j) => (j.date_found > max ? j.date_found : max), "");
  const sources = useMemo(() => [...new Set(jobs.map((j) => j.source))].sort(), [jobs]);

  const filtered = useMemo(() => {
    const c = norm(company.trim());
    const l = norm(location.trim());
    const rows = jobs.filter(
      (j) =>
        (!c || norm(j.company).includes(c)) &&
        (!l || norm(j.location).includes(l)) &&
        (!source || j.source === source) &&
        (!fromDate || j.date_found >= fromDate) &&
        (!toDate || j.date_found <= toDate)
    );
    const { key, dir } = sort;
    const mult = dir === "asc" ? 1 : -1;
    return [...rows].sort((a, b) => {
      const cmp = String(a[key]).localeCompare(String(b[key]), undefined, { sensitivity: "base" });
      return cmp !== 0 ? cmp * mult : b.date_found.localeCompare(a.date_found);
    });
  }, [jobs, company, location, source, fromDate, toDate, sort]);

  const stats = useMemo(() => {
    const companies = new Set(filtered.map((j) => j.company.toLowerCase()));
    const locCount = {};
    filtered.forEach((j) => (locCount[j.location] = (locCount[j.location] || 0) + 1));
    const topLoc = Object.entries(locCount).sort((a, b) => b[1] - a[1])[0];
    return {
      shown: filtered.length,
      newest: filtered.filter((j) => j.date_found === latestDate).length,
      companies: companies.size,
      topLocation: topLoc ? `${topLoc[0]} (${topLoc[1]})` : "-",
    };
  }, [filtered, latestDate]);

  function applyRange(r) {
    setActiveRange(r.label);
    setToDate("");
    setFromDate(r.days ? isoDaysAgo(r.days - 1) : "");
  }

  function resetFilters() {
    setCompany("");
    setLocation("");
    setSource("");
    setFromDate("");
    setToDate("");
    setActiveRange(null);
  }

  function toggleSort(key) {
    setSort((s) => ({ key, dir: s.key === key && s.dir === "desc" ? "asc" : "desc" }));
  }

  const sortMark = (key) => (sort.key === key ? (sort.dir === "asc" ? " ▲" : " ▼") : "");

  return (
    <div className="page">
      <header className="header">
        <div>
          <h1>Job Tracker</h1>
          <p className="muted">
            Senior Java Fullstack roles · Germany, Netherlands & EU Remote · last 30 days
          </p>
        </div>
        <div className="header-right">
          {data.generated_at && (
            <span className="muted small">
              Updated {new Date(data.generated_at).toLocaleString()}
            </span>
          )}
          <a className="btn" href={EXCEL_URL}>
            Download Excel
          </a>
        </div>
      </header>

      <section className="stats">
        <Stat label="Jobs shown" value={stats.shown} />
        <Stat label="Found on latest run" value={stats.newest} />
        <Stat label="Companies" value={stats.companies} />
        <Stat label="Top location" value={stats.topLocation} small />
      </section>

      <section className="filters">
        <div className="field grow">
          <label htmlFor="company">Company</label>
          <input
            id="company"
            type="search"
            placeholder="Search company…"
            value={company}
            onChange={(e) => setCompany(e.target.value)}
          />
        </div>
        <div className="field grow">
          <label htmlFor="location">Location</label>
          <input
            id="location"
            type="search"
            placeholder="e.g. Berlin, Remote…"
            value={location}
            onChange={(e) => setLocation(e.target.value)}
          />
        </div>
        <div className="field">
          <label htmlFor="source">Source</label>
          <select id="source" value={source} onChange={(e) => setSource(e.target.value)}>
            <option value="">All sources</option>
            {sources.map((s) => (
              <option key={s}>{s}</option>
            ))}
          </select>
        </div>
        <div className="field">
          <label htmlFor="from">From</label>
          <input
            id="from"
            type="date"
            value={fromDate}
            max={toDate || undefined}
            onChange={(e) => {
              setFromDate(e.target.value);
              setActiveRange(null);
            }}
          />
        </div>
        <div className="field">
          <label htmlFor="to">To</label>
          <input
            id="to"
            type="date"
            value={toDate}
            min={fromDate || undefined}
            onChange={(e) => {
              setToDate(e.target.value);
              setActiveRange(null);
            }}
          />
        </div>
      </section>

      <div className="chips">
        {QUICK_RANGES.map((r) => (
          <button
            key={r.label}
            className={`chip ${activeRange === r.label ? "active" : ""}`}
            onClick={() => applyRange(r)}
          >
            {r.label}
          </button>
        ))}
        <button className="chip ghost" onClick={resetFilters}>
          Reset filters
        </button>
      </div>

      {status === "loading" && <p className="empty">Loading jobs…</p>}
      {status === "error" && <p className="empty">Could not load jobs.json.</p>}
      {status === "ready" && filtered.length === 0 && (
        <p className="empty">
          {jobs.length === 0
            ? "No jobs yet. The tracker adds new jobs once a day."
            : "No jobs match these filters."}
        </p>
      )}

      {filtered.length > 0 && (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th onClick={() => toggleSort("date_found")}>Date{sortMark("date_found")}</th>
                <th onClick={() => toggleSort("company")}>Company{sortMark("company")}</th>
                <th onClick={() => toggleSort("title")}>Role{sortMark("title")}</th>
                <th onClick={() => toggleSort("location")}>Location{sortMark("location")}</th>
                <th>Tech stack</th>
                <th onClick={() => toggleSort("source")}>Source{sortMark("source")}</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((j, i) => (
                <tr key={`${j.contact_info}-${i}`}>
                  <td data-label="Date">
                    {formatDate(j.date_found)}
                    {j.date_found === latestDate && <span className="badge">New</span>}
                  </td>
                  <td data-label="Company" className="strong">{j.company}</td>
                  <td data-label="Role">{j.title}</td>
                  <td data-label="Location">{j.location}</td>
                  <td data-label="Tech stack">
                    <div className="tags">
                      {j.tech_stack
                        .split(",")
                        .map((t) => t.trim())
                        .filter(Boolean)
                        .map((t) => (
                          <span key={t} className="tag">{t}</span>
                        ))}
                    </div>
                  </td>
                  <td data-label="Source">{j.source}</td>
                  <td className="action">
                    {isLink(j.contact_info) ? (
                      <a className="btn small" href={j.contact_info} target="_blank" rel="noreferrer">
                        Apply ↗
                      </a>
                    ) : (
                      <span className="muted">{j.contact_info}</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function Stat({ label, value, small }) {
  return (
    <div className="stat">
      <div className={`stat-value ${small ? "small" : ""}`}>{value}</div>
      <div className="stat-label">{label}</div>
    </div>
  );
}
