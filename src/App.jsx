import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend
} from 'recharts';
import {
  LayoutGrid, FileText, PlusCircle, Search, Trash2, Pencil, X,
  AlertTriangle, CheckCircle2, Clock, FileEdit, ChevronRight,
  Paperclip, Image as ImageIcon, Download, File as FileIcon, Loader2,
  Users, LogOut, ShieldCheck, UserPlus, ClipboardList,
  ClipboardCheck, ArrowUp, ArrowDown, ListChecks, Ban, ShieldOff,
  Upload, Printer, Camera, FolderPlus, Layers, Copy
} from 'lucide-react';
import mammoth from 'mammoth';
import { supabase } from './supabaseClient';

const C = {
  bg: '#f4f6f8', surface: '#ffffff', border: '#dde3e8', borderStrong: '#c3ccd3',
  navy: '#12314a', navyDeep: '#0b2033', text: '#1c2b38', dim: '#647685', faint: '#94a3ae',
  green: '#2f9e6e', greenBg: '#e6f4ee', amber: '#c9862f', amberBg: '#fbf1e2',
  red: '#c0483d', redBg: '#fbeae8', blue: '#3b6e94', blueBg: '#e8f0f5',
};

const TYPE_OPTIONS = ['SOP', 'JSA', 'OPL', 'Inspection Checklist', 'Work Instruction', 'Training Record', 'Other'];
const STATUS_OPTIONS = ['Draft', 'Pending Approval', 'Approved'];
const MAX_FILE_BYTES = 20 * 1024 * 1024;
const FIELD_TYPES = ['Text', 'Notes', 'Number', 'Date', 'Yes/No', 'Pass/Fail', 'Multiple Choice', 'Photo'];
const CONDITIONABLE_TYPES = ['Yes/No', 'Pass/Fail', 'Multiple Choice'];
const ATTACH_BUCKET = 'attachments';

const STATUS_STYLE = {
  'Draft':            { fg: C.dim,   bg: '#eef1f3', icon: FileEdit },
  'Pending Approval':  { fg: C.amber, bg: C.amberBg, icon: Clock },
  'Approved':          { fg: C.green, bg: C.greenBg, icon: CheckCircle2 },
  'Expired':           { fg: C.red,   bg: C.redBg,   icon: AlertTriangle },
};

function uid(prefix) { return (prefix || 'x') + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8); }
function todayISO() { return new Date().toISOString().slice(0, 10); }
function daysUntil(dateStr) {
  if (!dateStr) return null;
  return Math.round((new Date(dateStr) - new Date(todayISO())) / (1000 * 60 * 60 * 24));
}
function effectiveStatus(doc) {
  if (doc.status === 'Approved' && doc.expiryDate) {
    const d = daysUntil(doc.expiryDate);
    if (d !== null && d < 0) return 'Expired';
  }
  return doc.status;
}
function fmtBytes(n) {
  if (n == null) return '';
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
  return (n / (1024 * 1024)).toFixed(1) + ' MB';
}
function titleFromFilename(name) {
  return name.replace(/\.docx?$/i, '').replace(/[_-]+/g, ' ').trim();
}
async function extractWordText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value.trim();
}
function optionsForField(field) {
  if (!field) return [];
  if (field.type === 'Yes/No') return ['Yes', 'No'];
  if (field.type === 'Pass/Fail') return ['Pass', 'Fail'];
  if (field.type === 'Multiple Choice') return (field.options || '').split(',').map(o => o.trim()).filter(Boolean);
  return [];
}
function fieldVisible(field, scopedValues) {
  if (!field.condition || !field.condition.fieldId) return true;
  return String(scopedValues[field.condition.fieldId] || '') === field.condition.equals;
}
// Builds a human-readable submission title from a template's titleTemplate string,
// e.g. "{{Building Number}} - {{Location}}" -> "12 - Brandon". Only pulls from
// non-repeating fields, since a repeating section can have many/no rows.
function interpolateTitle(template, values) {
  if (!template.titleTemplate || !template.titleTemplate.trim()) return '';
  const lookup = {};
  (template.sections || []).forEach(s => {
    if (!s.repeating) s.fields.forEach(f => { lookup[f.label.trim().toLowerCase()] = values[f.id]; });
  });
  let result = template.titleTemplate.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (match, label) => {
    const v = lookup[label.trim().toLowerCase()];
    return (v !== undefined && v !== null && v !== '') ? String(v) : '';
  });
  return result.replace(/\s{2,}/g, ' ').replace(/^[\s\-–—]+|[\s\-–—]+$/g, '').trim();
}
function Badge({ status }) {
  const s = STATUS_STYLE[status] || STATUS_STYLE['Draft'];
  const Icon = s.icon;
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: s.bg, color: s.fg, fontSize: 12, fontWeight: 600, padding: '4px 10px', borderRadius: 4, whiteSpace: 'nowrap' }}>
      <Icon size={12} strokeWidth={2.5} />{status}
    </span>
  );
}
function PassFailBadge({ value }) {
  const pass = (value || '').toLowerCase() === 'pass';
  const fail = (value || '').toLowerCase() === 'fail';
  const color = pass ? C.green : fail ? C.red : C.faint;
  const bg = pass ? C.greenBg : fail ? C.redBg : '#eef1f3';
  return <span style={{ fontSize: 12, fontWeight: 600, color, background: bg, padding: '3px 9px', borderRadius: 4 }}>{value || '—'}</span>;
}

// ---------- DB row <-> app object mapping ----------
function mapDoc(r) {
  return {
    id: r.id, title: r.title, type: r.type, department: r.department, owner: r.owner,
    revision: r.revision, status: r.status, effectiveDate: r.effective_date, expiryDate: r.expiry_date,
    content: r.content, attachments: r.attachments || [],
    createdBy: r.created_by_name, createdAt: r.created_at ? r.created_at.slice(0, 10) : null,
    updatedBy: r.updated_by_name, updatedAt: r.updated_at ? r.updated_at.slice(0, 10) : null,
  };
}
function mapTemplate(r) {
  let sections = r.sections;
  if (!sections || sections.length === 0) {
    // Backward compatibility with templates saved before sections existed.
    sections = (r.fields && r.fields.length > 0) ? [{ id: 'legacy', title: 'Fields', repeating: false, fields: r.fields }] : [];
  }
  return {
    id: r.id, name: r.name, category: r.category, department: r.department, sections,
    titleTemplate: r.title_template || '', description: r.description || '', attachments: r.attachments || [],
    createdBy: r.created_by_name, createdAt: r.created_at, updatedBy: r.updated_by_name, updatedAt: r.updated_at,
  };
}
function mapSubmission(r) {
  let sectionsSnapshot = r.sections_snapshot;
  if (!sectionsSnapshot || sectionsSnapshot.length === 0) {
    sectionsSnapshot = (r.fields_snapshot && r.fields_snapshot.length > 0) ? [{ id: 'legacy', title: 'Fields', repeating: false, fields: r.fields_snapshot }] : [];
  }
  return {
    id: r.id, templateId: r.template_id, templateName: r.template_name, category: r.category,
    department: r.department, sectionsSnapshot, values: r.values || {},
    title: r.title || r.template_name,
    filledBy: r.filled_by_name, filledAt: r.filled_at, filledAtTime: r.filled_at_time,
  };
}

const emptyForm = {
  id: null, title: '', type: 'SOP', department: '', owner: '',
  revision: 'A', status: 'Draft', effectiveDate: todayISO(), expiryDate: '',
  content: '', attachments: []
};
const emptyTemplateForm = { id: null, name: '', category: 'Inspection Checklist', department: '', sections: [], description: '', attachments: [], titleTemplate: '' };

export default function App() {
  const [authLoaded, setAuthLoaded] = useState(false);
  const [authUser, setAuthUser] = useState(null);
  const [profile, setProfile] = useState(null);
  const [authMode, setAuthMode] = useState('login');
  const [authError, setAuthError] = useState(null);
  const [authInfo, setAuthInfo] = useState(null);
  const [recoveryMode, setRecoveryMode] = useState(false);

  const [docs, setDocs] = useState([]);
  const [templates, setTemplates] = useState([]);
  const [submissions, setSubmissions] = useState([]);
  const [dataLoaded, setDataLoaded] = useState(false);
  const [saveError, setSaveError] = useState(null);
  const [saving, setSaving] = useState(false);
  const [view, setView] = useState('dashboard');
  const [form, setForm] = useState(emptyForm);
  const [templateForm, setTemplateForm] = useState(emptyTemplateForm);
  const [fillingTemplate, setFillingTemplate] = useState(null);
  const [removedAttachments, setRemovedAttachments] = useState([]);
  const [removedTemplateAttachments, setRemovedTemplateAttachments] = useState([]);
  const [search, setSearch] = useState('');
  const [filterType, setFilterType] = useState('All');
  const [filterStatus, setFilterStatus] = useState('All');
  const [activeDoc, setActiveDoc] = useState(null);
  const [activeSubmission, setActiveSubmission] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [confirmDeleteTemplate, setConfirmDeleteTemplate] = useState(null);
  const [attachmentUrlCache, setAttachmentUrlCache] = useState({});
  const [loadingAttachments, setLoadingAttachments] = useState({});

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setAuthUser(data.session?.user || null);
      setAuthLoaded(true);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((event, session) => {
      if (event === 'PASSWORD_RECOVERY') setRecoveryMode(true);
      setAuthUser(session?.user || null);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!authUser) { setProfile(null); return; }
    (async () => {
      const { data, error } = await supabase.from('profiles').select('*').eq('id', authUser.id).single();
      if (!error) setProfile(data);
    })();
  }, [authUser]);

  useEffect(() => {
    if (!authUser || !profile || !profile.active) return;
    (async () => {
      const [d, t, s] = await Promise.all([
        supabase.from('documents').select('*').order('created_at', { ascending: false }),
        supabase.from('templates').select('*').order('created_at', { ascending: false }),
        supabase.from('submissions').select('*').order('filled_at_time', { ascending: false }),
      ]);
      if (d.data) setDocs(d.data.map(mapDoc));
      if (t.data) setTemplates(t.data.map(mapTemplate));
      if (s.data) setSubmissions(s.data.map(mapSubmission));
      setDataLoaded(true);
    })();
  }, [authUser, profile]);

  async function refreshDocs() {
    const { data } = await supabase.from('documents').select('*').order('created_at', { ascending: false });
    if (data) setDocs(data.map(mapDoc));
  }
  async function refreshTemplates() {
    const { data } = await supabase.from('templates').select('*').order('created_at', { ascending: false });
    if (data) setTemplates(data.map(mapTemplate));
  }
  async function refreshSubmissions() {
    const { data } = await supabase.from('submissions').select('*').order('filled_at_time', { ascending: false });
    if (data) setSubmissions(data.map(mapSubmission));
  }

  function logout() { supabase.auth.signOut(); setView('dashboard'); }

  if (!authLoaded) return <CenteredMessage text="Loading…" />;
  if (recoveryMode) return <UpdatePasswordScreen onDone={() => setRecoveryMode(false)} />;

  if (!authUser) {
    return (
      <AuthGate
        mode={authMode} setMode={setAuthMode} error={authError} info={authInfo}
        onLogin={async (email, password) => {
          setAuthError(null); setAuthInfo(null);
          const { error } = await supabase.auth.signInWithPassword({ email, password });
          if (error) setAuthError(error.message);
        }}
        onSignup={async (name, email, password) => {
          setAuthError(null); setAuthInfo(null);
          const { data, error } = await supabase.auth.signUp({ email, password, options: { data: { name } } });
          if (error) { setAuthError(error.message); return; }
          if (!data.session) { setAuthInfo('Account created. Check your email to confirm it, then log in.'); setAuthMode('login'); }
        }}
        onForgotPassword={async (email) => {
          setAuthError(null); setAuthInfo(null);
          const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin });
          if (error) setAuthError(error.message); else setAuthInfo('Password reset email sent — check your inbox.');
        }}
      />
    );
  }

  if (!profile) return <CenteredMessage text="Setting up your account…" />;

  if (!profile.active) {
    return (
      <AuthShell>
        <ShieldOff size={22} color={C.red} style={{ marginBottom: 10 }} />
        <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, margin: '0 0 8px' }}>Access disabled</h1>
        <p style={{ fontSize: 13.5, color: C.dim, marginBottom: 20, lineHeight: 1.6 }}>Your account has been deactivated. Contact your workspace admin if you think this is a mistake.</p>
        <button onClick={logout} style={{ ...btnGhost(C), width: '100%', justifyContent: 'center' }}>Sign Out</button>
      </AuthShell>
    );
  }

  const isAdmin = profile.role === 'admin';
  const canBuildTemplates = isAdmin || !!profile.can_build;
  const canBuildDocuments = isAdmin || !!profile.can_build_docs;
  const displayName = profile.name || authUser.email;

  function startNew() { setForm({ ...emptyForm, id: null, attachments: [] }); setRemovedAttachments([]); setView('form'); }
  function startEdit(doc) { setForm({ ...doc, attachments: doc.attachments || [] }); setRemovedAttachments([]); setView('form'); }

  async function saveForm(e) {
    e.preventDefault();
    if (!form.title.trim()) return;
    setSaving(true); setSaveError(null);
    try {
      for (const path of removedAttachments) { await supabase.storage.from(ATTACH_BUCKET).remove([path]).catch(() => {}); }
      const cleaned = [];
      for (const a of (form.attachments || [])) {
        if (a.isNew && a.file) {
          const path = `${uid('att')}-${a.file.name}`;
          const { error: upErr } = await supabase.storage.from(ATTACH_BUCKET).upload(path, a.file);
          if (upErr) { setSaveError(`"${a.name}" didn't upload: ${upErr.message}`); continue; }
          cleaned.push({ id: a.id, name: a.name, type: a.type, size: a.size, path });
        } else { cleaned.push({ id: a.id, name: a.name, type: a.type, size: a.size, path: a.path }); }
      }
      const payload = {
        title: form.title, type: form.type, department: form.department || null, owner: form.owner || null,
        revision: form.revision || 'A', status: form.status,
        effective_date: form.effectiveDate || null, expiry_date: form.expiryDate || null,
        content: form.content || null, attachments: cleaned,
        updated_by_name: displayName, updated_at: new Date().toISOString(),
      };
      if (form.id) {
        const { error } = await supabase.from('documents').update(payload).eq('id', form.id);
        if (error) throw error;
      } else {
        const { error } = await supabase.from('documents').insert({ ...payload, created_by: authUser.id, created_by_name: displayName });
        if (error) throw error;
      }
      await refreshDocs();
      setView('documents');
    } catch (err) {
      setSaveError(err.message || 'Could not save this document.');
    } finally { setSaving(false); }
  }

  async function deleteDoc(id) {
    const doc = docs.find(d => d.id === id);
    const { error } = await supabase.from('documents').delete().eq('id', id);
    if (!error) {
      for (const a of (doc?.attachments || [])) { if (a.path) await supabase.storage.from(ATTACH_BUCKET).remove([a.path]).catch(() => {}); }
      await refreshDocs();
    } else { setSaveError(error.message); }
    setConfirmDelete(null);
    if (activeDoc && activeDoc.id === id) setActiveDoc(null);
  }

  async function quickStatus(id, status) {
    const { error } = await supabase.from('documents').update({ status, updated_by_name: displayName, updated_at: new Date().toISOString() }).eq('id', id);
    if (error) setSaveError(error.message); else refreshDocs();
  }

  async function loadAttachment(att) {
    const key = att.id;
    if (attachmentUrlCache[key] || loadingAttachments[key] || !att.path) return;
    setLoadingAttachments(s => ({ ...s, [key]: true }));
    try {
      const { data, error } = await supabase.storage.from(ATTACH_BUCKET).download(att.path);
      if (!error && data) setAttachmentUrlCache(c => ({ ...c, [key]: URL.createObjectURL(data) }));
    } finally { setLoadingAttachments(s => ({ ...s, [key]: false })); }
  }

  // ---------- Templates ----------
  function startNewTemplate() { setTemplateForm({ ...emptyTemplateForm, id: null, sections: [], attachments: [] }); setRemovedTemplateAttachments([]); setView('templateForm'); }
  function startEditTemplate(t) { setTemplateForm({ ...t, attachments: t.attachments || [] }); setRemovedTemplateAttachments([]); setView('templateForm'); }

  async function saveTemplate(e) {
    e.preventDefault();
    if (!templateForm.name.trim() || templateForm.sections.length === 0) return;
    setSaving(true); setSaveError(null);
    try {
      for (const path of removedTemplateAttachments) { await supabase.storage.from(ATTACH_BUCKET).remove([path]).catch(() => {}); }
      const cleanedAttachments = [];
      for (const a of (templateForm.attachments || [])) {
        if (a.isNew && a.file) {
          const path = `${uid('att')}-${a.file.name}`;
          const { error: upErr } = await supabase.storage.from(ATTACH_BUCKET).upload(path, a.file);
          if (upErr) { setSaveError(`"${a.name}" didn't upload: ${upErr.message}`); continue; }
          cleanedAttachments.push({ id: a.id, name: a.name, type: a.type, size: a.size, path });
        } else { cleanedAttachments.push({ id: a.id, name: a.name, type: a.type, size: a.size, path: a.path }); }
      }
      const payload = {
        name: templateForm.name, category: templateForm.category, department: templateForm.department || null,
        sections: templateForm.sections, title_template: templateForm.titleTemplate || null,
        description: templateForm.description || null, attachments: cleanedAttachments,
        updated_by_name: displayName, updated_at: new Date().toISOString(),
      };
      if (templateForm.id) {
        const { error } = await supabase.from('templates').update(payload).eq('id', templateForm.id);
        if (error) throw error;
      } else {
        const { error } = await supabase.from('templates').insert({ ...payload, created_by_name: displayName });
        if (error) throw error;
      }
      await refreshTemplates();
      setView('templates');
    } catch (err) {
      setSaveError(err.message || 'Could not save this template.');
    } finally { setSaving(false); }
  }

  async function deleteTemplate(id) {
    const t = templates.find(x => x.id === id);
    const { error } = await supabase.from('templates').delete().eq('id', id);
    if (!error) {
      for (const a of (t?.attachments || [])) { if (a.path) await supabase.storage.from(ATTACH_BUCKET).remove([a.path]).catch(() => {}); }
      await refreshTemplates();
    } else { setSaveError(error.message); }
    setConfirmDeleteTemplate(null);
  }

  function startFill(t) { setFillingTemplate(t); setView('fillForm'); }

  async function submitFilledForm(rawValues) {
    setSaving(true); setSaveError(null);
    try {
      const finalValues = {};
      for (const section of fillingTemplate.sections) {
        if (section.repeating) {
          const rows = rawValues[section.id] || [];
          const finalRows = [];
          for (const row of rows) {
            const finalRow = {};
            for (const f of section.fields) {
              const v = row[f.id];
              if (f.type === 'Photo' && v instanceof File) {
                const path = `${uid('att')}-${v.name}`;
                const { error: upErr } = await supabase.storage.from(ATTACH_BUCKET).upload(path, v);
                if (upErr) { setSaveError(`"${v.name}" didn't upload: ${upErr.message}`); continue; }
                finalRow[f.id] = { id: uid('sp'), name: v.name, type: v.type, size: v.size, path };
              } else if (v !== undefined) { finalRow[f.id] = v; }
            }
            finalRows.push(finalRow);
          }
          finalValues[section.id] = finalRows;
        } else {
          for (const f of section.fields) {
            const v = rawValues[f.id];
            if (f.type === 'Photo' && v instanceof File) {
              const path = `${uid('att')}-${v.name}`;
              const { error: upErr } = await supabase.storage.from(ATTACH_BUCKET).upload(path, v);
              if (upErr) { setSaveError(`"${v.name}" didn't upload: ${upErr.message}`); continue; }
              finalValues[f.id] = { id: uid('sp'), name: v.name, type: v.type, size: v.size, path };
            } else if (v !== undefined) { finalValues[f.id] = v; }
          }
        }
      }
      const title = interpolateTitle(fillingTemplate, finalValues);
      const { error } = await supabase.from('submissions').insert({
        template_id: fillingTemplate.id, template_name: fillingTemplate.name, category: fillingTemplate.category,
        department: fillingTemplate.department, sections_snapshot: fillingTemplate.sections, values: finalValues,
        title: title || null, filled_by_name: displayName, filled_at: todayISO(),
      });
      if (error) throw error;
      await refreshSubmissions();
      setFillingTemplate(null);
      setView('submissions');
    } catch (err) {
      setSaveError(err.message || 'Could not save this entry.');
    } finally { setSaving(false); }
  }

  // ---------- Derived ----------
  const withStatus = useMemo(() => docs.map(d => ({ ...d, effStatus: effectiveStatus(d) })), [docs]);
  const filtered = useMemo(() => withStatus.filter(d => {
    if (filterType !== 'All' && d.type !== filterType) return false;
    if (filterStatus !== 'All' && d.effStatus !== filterStatus) return false;
    if (search.trim()) {
      const q = search.toLowerCase();
      if (!d.title.toLowerCase().includes(q) && !(d.department || '').toLowerCase().includes(q) && !(d.owner || '').toLowerCase().includes(q)) return false;
    }
    return true;
  }), [withStatus, filterType, filterStatus, search]);

  const kpis = useMemo(() => {
    const total = withStatus.length;
    const approved = withStatus.filter(d => d.effStatus === 'Approved').length;
    const pending = withStatus.filter(d => d.effStatus === 'Pending Approval').length;
    const expired = withStatus.filter(d => d.effStatus === 'Expired').length;
    const expiringSoon = withStatus.filter(d => {
      if (d.effStatus !== 'Approved' || !d.expiryDate) return false;
      const days = daysUntil(d.expiryDate);
      return days !== null && days >= 0 && days <= 30;
    }).length;
    const weekAgo = new Date(); weekAgo.setDate(weekAgo.getDate() - 7);
    const submissionsThisWeek = submissions.filter(s => new Date(s.filledAtTime || s.filledAt) >= weekAgo).length;
    return { total, approved, pending, expired, expiringSoon, submissionsThisWeek };
  }, [withStatus, submissions]);

  const statusChartData = useMemo(() => {
    const counts = { Approved: 0, 'Pending Approval': 0, Draft: 0, Expired: 0 };
    withStatus.forEach(d => { counts[d.effStatus] = (counts[d.effStatus] || 0) + 1; });
    return Object.entries(counts).filter(([, v]) => v > 0).map(([name, value]) => ({ name, value }));
  }, [withStatus]);
  const typeChartData = useMemo(() => {
    const counts = {};
    withStatus.forEach(d => { counts[d.type] = (counts[d.type] || 0) + 1; });
    return Object.entries(counts).map(([name, count]) => ({ name, count }));
  }, [withStatus]);
  const expiringList = useMemo(() => withStatus
    .filter(d => d.effStatus === 'Approved' && d.expiryDate)
    .map(d => ({ ...d, days: daysUntil(d.expiryDate) }))
    .filter(d => d.days !== null && d.days <= 30)
    .sort((a, b) => a.days - b.days).slice(0, 6), [withStatus]);
  const recentSubmissions = useMemo(() =>
    [...submissions].sort((a, b) => new Date(b.filledAtTime || b.filledAt) - new Date(a.filledAtTime || a.filledAt)).slice(0, 6),
    [submissions]);

  const NavItem = ({ id, icon: Icon, label }) => (
    <button onClick={() => setView(id)} style={{
      display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '10px 14px', borderRadius: 6, border: 'none', cursor: 'pointer',
      background: view === id ? 'rgba(255,255,255,0.1)' : 'transparent', color: view === id ? '#fff' : 'rgba(255,255,255,0.65)',
      fontSize: 14, fontWeight: 500, textAlign: 'left', fontFamily: 'inherit'
    }}>
      <Icon size={16} />{label}
    </button>
  );

  const inputStyle = { width: '100%', padding: '9px 12px', borderRadius: 5, border: `1px solid ${C.border}`, fontSize: 14, color: C.text, background: '#fff', fontFamily: 'inherit' };
  const labelStyle = { fontSize: 12, fontWeight: 600, color: C.dim, marginBottom: 6, display: 'block', textTransform: 'uppercase', letterSpacing: '0.03em' };

  if (!dataLoaded) return <CenteredMessage text="Loading your workspace…" />;

  return (
    <div style={{ display: 'flex', minHeight: '100vh', background: C.bg, fontFamily: "'Inter', -apple-system, sans-serif", color: C.text }}>
      <div style={{ width: 220, background: C.navyDeep, padding: '22px 14px', display: 'flex', flexDirection: 'column', flexShrink: 0 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, color: '#fff', padding: '0 10px 24px 10px', letterSpacing: '0.01em' }}>
          Production<br /><span style={{ color: '#7fc7a4' }}>Document Center</span>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          <NavItem id="dashboard" icon={LayoutGrid} label="Dashboard" />
          <NavItem id="documents" icon={FileText} label="Documents" />
          <NavItem id="templates" icon={ClipboardList} label="Templates" />
          <NavItem id="submissions" icon={ListChecks} label="Submissions" />
          {isAdmin && <NavItem id="users" icon={Users} label="Users" />}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 20 }}>
          {canBuildDocuments && (
            <button onClick={startNew} style={{ display: 'flex', alignItems: 'center', gap: 8, background: C.green, color: '#fff', border: 'none', borderRadius: 6, padding: '10px 14px', fontSize: 13.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit' }}>
              <PlusCircle size={16} /> New Document
            </button>
          )}
          {canBuildTemplates && (
            <button onClick={startNewTemplate} style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'rgba(255,255,255,0.08)', color: '#fff', border: `1px solid rgba(255,255,255,0.15)`, borderRadius: 6, padding: '10px 14px', fontSize: 13.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit' }}>
              <ClipboardList size={16} /> New Template
            </button>
          )}
        </div>

        <div style={{ marginTop: 'auto', paddingTop: 20, borderTop: '1px solid rgba(255,255,255,0.1)' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
            <div style={{ width: 28, height: 28, borderRadius: '50%', background: isAdmin ? C.green : C.blue, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 12, fontWeight: 700, flexShrink: 0 }}>
              {displayName.slice(0, 1).toUpperCase()}
            </div>
            <div style={{ overflow: 'hidden' }}>
              <div style={{ color: '#fff', fontSize: 13, fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{displayName}</div>
              <div style={{ color: 'rgba(255,255,255,0.5)', fontSize: 11, display: 'flex', alignItems: 'center', gap: 3 }}>
                {isAdmin && <ShieldCheck size={11} />} {isAdmin ? 'Admin' : 'Member'}
              </div>
            </div>
          </div>
          <button onClick={logout} style={{ display: 'flex', alignItems: 'center', gap: 6, width: '100%', background: 'none', border: 'none', color: 'rgba(255,255,255,0.55)', fontSize: 12.5, cursor: 'pointer', padding: '6px 4px', fontFamily: 'inherit' }}>
            <LogOut size={13} /> Log out
          </button>
        </div>
      </div>

      <div style={{ flex: 1, padding: '28px 32px', overflow: 'auto' }}>
        {saveError && (
          <div style={{ marginBottom: 18, padding: '10px 14px', borderRadius: 6, background: C.redBg, color: C.red, fontSize: 13, fontWeight: 500 }}>{saveError}</div>
        )}

        {view === 'dashboard' && (
          <Dashboard kpis={kpis} statusChartData={statusChartData} typeChartData={typeChartData} expiringList={expiringList}
            recentSubmissions={recentSubmissions}
            onSelectDoc={(d) => { setActiveDoc(d); setView('documents'); }}
            onSelectSubmission={(s) => { setActiveSubmission(s); setView('submissions'); }}
            total={docs.length} onNew={startNew} canBuild={canBuildDocuments} />
        )}

        {view === 'documents' && (
          <DocumentsList filtered={filtered} search={search} setSearch={setSearch} filterType={filterType} setFilterType={setFilterType}
            filterStatus={filterStatus} setFilterStatus={setFilterStatus} onEdit={startEdit}
            onDelete={(id) => setConfirmDelete(id)} onView={(d) => setActiveDoc(d)} onQuickStatus={quickStatus}
            hasAny={docs.length > 0} onNew={startNew} isAdmin={isAdmin} canBuild={canBuildDocuments} />
        )}

        {view === 'form' && canBuildDocuments && (
          <DocumentForm form={form} setForm={setForm} onCancel={() => setView(form.id ? 'documents' : 'dashboard')}
            onSubmit={saveForm} onRemoveExisting={(path) => setRemovedAttachments(p => [...p, path])}
            inputStyle={inputStyle} labelStyle={labelStyle} saving={saving} />
        )}

        {view === 'templates' && (
          <TemplatesList templates={templates} isAdmin={isAdmin} canBuild={canBuildTemplates} onNew={startNewTemplate} onEdit={startEditTemplate}
            onDelete={(id) => setConfirmDeleteTemplate(id)} onFill={startFill} submissions={submissions} />
        )}

        {view === 'templateForm' && canBuildTemplates && (
          <TemplateForm templateForm={templateForm} setTemplateForm={setTemplateForm} onCancel={() => setView('templates')}
            onSubmit={saveTemplate} onRemoveExisting={(path) => setRemovedTemplateAttachments(p => [...p, path])}
            inputStyle={inputStyle} labelStyle={labelStyle} saving={saving} />
        )}

        {view === 'fillForm' && fillingTemplate && (
          <FillForm template={fillingTemplate} onCancel={() => { setFillingTemplate(null); setView('templates'); }}
            onSubmit={submitFilledForm} inputStyle={inputStyle} labelStyle={labelStyle} saving={saving} sessionName={displayName} />
        )}

        {view === 'submissions' && (
          <SubmissionsList submissions={submissions} templates={templates} onView={(s) => setActiveSubmission(s)} />
        )}

        {view === 'users' && isAdmin && (
          <UserManagement currentUserId={authUser.id} onUpdateProfile={async (id, patch) => {
            const { error } = await supabase.from('profiles').update(patch).eq('id', id);
            if (error) setAuthError(error.message);
          }} />
        )}
      </div>

      {activeDoc && (
        <Modal onClose={() => setActiveDoc(null)}>
          <div id="print-area">
            <div style={{ fontSize: 10.5, color: C.faint, marginBottom: 14, fontWeight: 600, letterSpacing: '0.04em', textTransform: 'uppercase' }}>Production Document Center</div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4 }}>
              <div>
                <div style={{ fontSize: 11, color: C.faint, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 6 }}>{activeDoc.type} · REV. {activeDoc.revision || 'A'}</div>
                <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>{activeDoc.title}</h2>
              </div>
              <Badge status={effectiveStatus(activeDoc)} />
            </div>
            <div style={{ display: 'flex', gap: 24, margin: '18px 0', fontSize: 13, color: C.dim, flexWrap: 'wrap' }}>
              <span><strong style={{ color: C.text }}>Department:</strong> {activeDoc.department || '—'}</span>
              <span><strong style={{ color: C.text }}>Owner:</strong> {activeDoc.owner || '—'}</span>
              <span><strong style={{ color: C.text }}>Effective:</strong> {activeDoc.effectiveDate || '—'}</span>
              <span><strong style={{ color: C.text }}>Expires:</strong> {activeDoc.expiryDate || '—'}</span>
            </div>
            {(activeDoc.createdBy || activeDoc.updatedBy) && (
              <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 14 }}>
                {activeDoc.createdBy && <>Created by {activeDoc.createdBy}{activeDoc.createdAt ? ` on ${activeDoc.createdAt}` : ''}. </>}
                {activeDoc.updatedBy && <>Last edited by {activeDoc.updatedBy}{activeDoc.updatedAt ? ` on ${activeDoc.updatedAt}` : ''}.</>}
              </div>
            )}
            <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 6, padding: 16, fontSize: 14, lineHeight: 1.7, whiteSpace: 'pre-wrap', minHeight: 100 }}>
              {activeDoc.content ? activeDoc.content : <span style={{ color: C.faint }}>No content added yet.</span>}
            </div>
            {(activeDoc.attachments && activeDoc.attachments.length > 0) && (
              <div style={{ marginTop: 20 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: C.dim, textTransform: 'uppercase', letterSpacing: '0.03em', marginBottom: 10 }}>Attachments ({activeDoc.attachments.length})</div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(140px, 1fr))', gap: 10 }}>
                  {activeDoc.attachments.map(a => (
                    <AttachmentTile key={a.id} meta={a} url={attachmentUrlCache[a.id]} isLoading={!!loadingAttachments[a.id]} onLoad={() => loadAttachment(a)} />
                  ))}
                </div>
              </div>
            )}
            <div style={{ marginTop: 24, paddingTop: 12, borderTop: `1px solid ${C.border}`, fontSize: 11, color: C.faint }}>
              Printed {todayISO()}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
            {canBuildDocuments && <button onClick={() => { setActiveDoc(null); startEdit(activeDoc); }} style={btnPrimary(C)}><Pencil size={14} /> Edit Document</button>}
            <button onClick={() => window.print()} style={btnGhost(C)}><Printer size={14} /> Export PDF</button>
            <button onClick={() => setActiveDoc(null)} style={btnGhost(C)}>Close</button>
          </div>
        </Modal>
      )}

      {activeSubmission && (
        <Modal onClose={() => setActiveSubmission(null)}>
          <div id="print-area">
            <div style={{ fontSize: 10.5, color: C.faint, marginBottom: 14, fontWeight: 600, letterSpacing: '0.04em', textTransform: 'uppercase' }}>Production Document Center</div>
            <div style={{ fontSize: 11, color: C.faint, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 6 }}>{activeSubmission.category}</div>
            <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 21, margin: '0 0 6px' }}>{activeSubmission.title || activeSubmission.templateName}</h2>
            {activeSubmission.title && activeSubmission.title !== activeSubmission.templateName && (
              <div style={{ fontSize: 12, color: C.faint, marginBottom: 6 }}>{activeSubmission.templateName}</div>
            )}
            <div style={{ fontSize: 12.5, color: C.dim, marginBottom: 18 }}>
              Filled out by <strong style={{ color: C.text }}>{activeSubmission.filledBy}</strong> on {activeSubmission.filledAt}
              {activeSubmission.department ? ` · ${activeSubmission.department}` : ''}
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 22 }}>
              {(activeSubmission.sectionsSnapshot || []).map(section => (
                <div key={section.id}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.navy, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${C.border}` }}>
                    {section.title || 'Fields'}
                  </div>
                  {section.repeating ? (
                    (activeSubmission.values[section.id] || []).length === 0 ? (
                      <div style={{ fontSize: 12.5, color: C.faint }}>No entries logged.</div>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                        {(activeSubmission.values[section.id] || []).map((row, i) => (
                          <div key={i} style={{ border: `1px solid ${C.border}`, borderRadius: 6, padding: 14 }}>
                            <div style={{ fontSize: 10.5, fontWeight: 700, color: C.faint, marginBottom: 10, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Entry {i + 1}</div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                              {section.fields.filter(f => fieldVisible(f, row)).map(f => (
                                <FieldAnswer key={f.id} field={f} value={row[f.id]} attachmentUrlCache={attachmentUrlCache} loadingAttachments={loadingAttachments} onLoadAttachment={loadAttachment} />
                              ))}
                            </div>
                          </div>
                        ))}
                      </div>
                    )
                  ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                      {section.fields.filter(f => fieldVisible(f, activeSubmission.values)).map(f => (
                        <FieldAnswer key={f.id} field={f} value={activeSubmission.values[f.id]} attachmentUrlCache={attachmentUrlCache} loadingAttachments={loadingAttachments} onLoadAttachment={loadAttachment} />
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div style={{ marginTop: 24, paddingTop: 12, borderTop: `1px solid ${C.border}`, fontSize: 11, color: C.faint }}>
              Printed {todayISO()}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 10, marginTop: 22 }}>
            <button onClick={() => window.print()} style={btnGhost(C)}><Printer size={14} /> Export PDF</button>
            <button onClick={() => setActiveSubmission(null)} style={btnGhost(C)}>Close</button>
          </div>
        </Modal>
      )}

      {confirmDelete && (
        <Modal onClose={() => setConfirmDelete(null)} narrow>
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
            <AlertTriangle size={20} color={C.red} style={{ flexShrink: 0, marginTop: 2 }} />
            <div>
              <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 6 }}>Delete this document?</div>
              <div style={{ fontSize: 13.5, color: C.dim }}>This removes it — and any attachments — for everyone. This can't be undone.</div>
            </div>
          </div>
          <div style={{ display: 'flex', gap: 10, marginTop: 22, justifyContent: 'flex-end' }}>
            <button onClick={() => setConfirmDelete(null)} style={btnGhost(C)}>Cancel</button>
            <button onClick={() => deleteDoc(confirmDelete)} style={{ ...btnPrimary(C), background: C.red }}><Trash2 size={14} /> Delete</button>
          </div>
        </Modal>
      )}

      {confirmDeleteTemplate && (
        <Modal onClose={() => setConfirmDeleteTemplate(null)} narrow>
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
            <AlertTriangle size={20} color={C.red} style={{ flexShrink: 0, marginTop: 2 }} />
            <div>
              <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 6 }}>Delete this template?</div>
              <div style={{ fontSize: 13.5, color: C.dim }}>People won't be able to fill it out anymore. Submissions already recorded from it are kept.</div>
            </div>
          </div>
          <div style={{ display: 'flex', gap: 10, marginTop: 22, justifyContent: 'flex-end' }}>
            <button onClick={() => setConfirmDeleteTemplate(null)} style={btnGhost(C)}>Cancel</button>
            <button onClick={() => deleteTemplate(confirmDeleteTemplate)} style={{ ...btnPrimary(C), background: C.red }}><Trash2 size={14} /> Delete</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ================= Auth screens =================
function CenteredMessage({ text }) {
  return <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', color: C.dim, fontFamily: "'Inter', sans-serif", fontSize: 14 }}>{text}</div>;
}
function AuthShell({ children }) {
  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: C.navyDeep, fontFamily: "'Inter', -apple-system, sans-serif", padding: 20 }}>
      <div style={{ background: '#fff', borderRadius: 12, padding: 36, width: '100%', maxWidth: 380, boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }}>{children}</div>
    </div>
  );
}
function AuthGate({ mode, setMode, onLogin, onSignup, onForgotPassword, error, info }) {
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [forgotOpen, setForgotOpen] = useState(false);
  const [forgotEmail, setForgotEmail] = useState('');

  function submit(e) {
    e.preventDefault();
    if (mode === 'signup') onSignup(name, email, password);
    else onLogin(email, password);
  }

  if (forgotOpen) {
    return (
      <AuthShell>
        <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, margin: '0 0 8px' }}>Reset your password</h1>
        <p style={{ fontSize: 13.5, color: C.dim, marginBottom: 18 }}>We'll email you a link to set a new password.</p>
        <label style={{ fontSize: 12, fontWeight: 600, color: C.dim, marginBottom: 6, display: 'block' }}>Email</label>
        <input autoFocus value={forgotEmail} onChange={e => setForgotEmail(e.target.value)} type="email"
          style={{ width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${C.border}`, fontSize: 14, marginBottom: 16, fontFamily: 'inherit' }} />
        {error && <div style={{ color: C.red, fontSize: 12.5, marginBottom: 14 }}>{error}</div>}
        {info && <div style={{ color: C.green, fontSize: 12.5, marginBottom: 14 }}>{info}</div>}
        <button onClick={() => onForgotPassword(forgotEmail)} style={{ ...btnPrimary(C), width: '100%', justifyContent: 'center', padding: '11px 18px', marginBottom: 10 }}>Send Reset Link</button>
        <button onClick={() => setForgotOpen(false)} style={{ ...btnGhost(C), width: '100%', justifyContent: 'center' }}>Back to Log In</button>
      </AuthShell>
    );
  }

  return (
    <AuthShell>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: '0 0 4px', color: C.navy }}>Production Document Center</h1>
      <p style={{ fontSize: 13.5, color: C.dim, margin: '0 0 22px' }}>{mode === 'signup' ? 'Create your account.' : 'Log in to your workspace.'}</p>
      <form onSubmit={submit}>
        {mode === 'signup' && (
          <>
            <label style={{ fontSize: 12, fontWeight: 600, color: C.dim, marginBottom: 6, display: 'block' }}>Your name</label>
            <input autoFocus value={name} onChange={e => setName(e.target.value)} placeholder="e.g. Felix"
              style={{ width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${C.border}`, fontSize: 14, marginBottom: 16, fontFamily: 'inherit' }} />
          </>
        )}
        <label style={{ fontSize: 12, fontWeight: 600, color: C.dim, marginBottom: 6, display: 'block' }}>Email</label>
        <input value={email} onChange={e => setEmail(e.target.value)} type="email"
          style={{ width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${C.border}`, fontSize: 14, marginBottom: 16, fontFamily: 'inherit' }} />
        <label style={{ fontSize: 12, fontWeight: 600, color: C.dim, marginBottom: 6, display: 'block' }}>Password</label>
        <input value={password} onChange={e => setPassword(e.target.value)} type="password"
          style={{ width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${C.border}`, fontSize: 14, marginBottom: 8, fontFamily: 'inherit' }} />
        {mode === 'login' && (
          <button type="button" onClick={() => { setForgotOpen(true); setForgotEmail(email); }} style={{ border: 'none', background: 'none', color: C.blue, fontSize: 12, cursor: 'pointer', padding: 0, marginBottom: 14 }}>Forgot password?</button>
        )}
        {mode === 'signup' && <div style={{ marginBottom: 14 }} />}
        {error && <div style={{ color: C.red, fontSize: 12.5, marginBottom: 14 }}>{error}</div>}
        {info && <div style={{ color: C.green, fontSize: 12.5, marginBottom: 14 }}>{info}</div>}
        <button type="submit" style={{ ...btnPrimary(C), width: '100%', justifyContent: 'center', padding: '11px 18px' }}>{mode === 'signup' ? 'Create Account' : 'Log In'}</button>
      </form>
      <div style={{ fontSize: 12.5, color: C.dim, marginTop: 18, textAlign: 'center' }}>
        {mode === 'signup' ? (
          <>Already have an account? <button onClick={() => setMode('login')} style={{ border: 'none', background: 'none', color: C.blue, cursor: 'pointer', fontSize: 12.5, padding: 0 }}>Log in</button></>
        ) : (
          <>New here? <button onClick={() => setMode('signup')} style={{ border: 'none', background: 'none', color: C.blue, cursor: 'pointer', fontSize: 12.5, padding: 0 }}>Create an account</button></>
        )}
      </div>
    </AuthShell>
  );
}
function UpdatePasswordScreen({ onDone }) {
  const [password, setPassword] = useState('');
  const [err, setErr] = useState(null);
  const [done, setDone] = useState(false);
  async function submit(e) {
    e.preventDefault();
    if (password.length < 6) { setErr('Password needs to be at least 6 characters.'); return; }
    const { error } = await supabase.auth.updateUser({ password });
    if (error) { setErr(error.message); return; }
    setDone(true);
    setTimeout(onDone, 1200);
  }
  return (
    <AuthShell>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, margin: '0 0 8px' }}>Set a new password</h1>
      {done ? (
        <p style={{ fontSize: 13.5, color: C.green }}>Password updated. Taking you back in…</p>
      ) : (
        <form onSubmit={submit}>
          <label style={{ fontSize: 12, fontWeight: 600, color: C.dim, marginBottom: 6, display: 'block' }}>New password</label>
          <input autoFocus value={password} onChange={e => setPassword(e.target.value)} type="password"
            style={{ width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${C.border}`, fontSize: 14, marginBottom: 16, fontFamily: 'inherit' }} />
          {err && <div style={{ color: C.red, fontSize: 12.5, marginBottom: 14 }}>{err}</div>}
          <button type="submit" style={{ ...btnPrimary(C), width: '100%', justifyContent: 'center', padding: '11px 18px' }}>Update Password</button>
        </form>
      )}
    </AuthShell>
  );
}

// ================= User Management =================
function UserManagement({ currentUserId, onUpdateProfile }) {
  const [users, setUsers] = useState(null);
  const [err, setErr] = useState(null);
  async function load() {
    const { data, error } = await supabase.from('profiles').select('*').order('created_at', { ascending: true });
    if (error) setErr(error.message); else setUsers(data);
  }
  useEffect(() => { load(); }, []);
  async function patch(id, fields) { await onUpdateProfile(id, fields); load(); }
  function toggleRole(u) {
    const admins = (users || []).filter(x => x.role === 'admin' && x.active);
    if (u.id === currentUserId) return;
    if (u.role === 'admin' && admins.length <= 1) return;
    patch(u.id, { role: u.role === 'admin' ? 'member' : 'admin' });
  }
  function toggleBuild(u) { patch(u.id, { can_build: !u.can_build }); }
  function toggleBuildDocs(u) { patch(u.id, { can_build_docs: !u.can_build_docs }); }
  function toggleActive(u) {
    if (u.id === currentUserId) return;
    const admins = (users || []).filter(x => x.role === 'admin' && x.active);
    if (u.role === 'admin' && u.active && admins.length <= 1) return;
    patch(u.id, { active: !u.active });
  }
  if (!users) return <CenteredMessage text="Loading users…" />;
  return (
    <div>
      <PageHeader title="Users" subtitle={`${users.length} ${users.length === 1 ? 'person has' : 'people have'} an account. People create their own accounts from the sign-up screen — grant permissions here afterward.`} />
      {err && <div style={{ color: C.red, fontSize: 13, marginBottom: 14 }}>{err}</div>}
      <div style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'hidden', maxWidth: 960 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1.1fr 0.8fr 0.95fr 0.95fr 0.9fr', padding: '10px 16px', fontSize: 11, fontWeight: 600, color: C.faint, textTransform: 'uppercase', letterSpacing: '0.04em', borderBottom: `1px solid ${C.border}`, background: C.bg }}>
          <span>Name</span><span>Role</span><span>Doc Builder</span><span>Template Builder</span><span>Access</span>
        </div>
        {users.map(u => (
          <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '1.1fr 0.8fr 0.95fr 0.95fr 0.9fr', padding: '12px 16px', fontSize: 13.5, borderBottom: `1px solid ${C.border}`, alignItems: 'center', opacity: u.active ? 1 : 0.55 }}>
            <span style={{ fontWeight: 600 }}>{u.name}{u.id === currentUserId ? ' (you)' : ''}</span>
            <button onClick={() => toggleRole(u)} disabled={u.id === currentUserId} style={{ border: 'none', background: 'none', cursor: u.id === currentUserId ? 'default' : 'pointer', padding: 0, textAlign: 'left' }}>
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, fontSize: 11.5, fontWeight: 600, color: u.role === 'admin' ? C.green : C.dim, background: u.role === 'admin' ? C.greenBg : '#eef1f3', padding: '3px 8px', borderRadius: 4 }}>
                {u.role === 'admin' && <ShieldCheck size={11} />}{u.role === 'admin' ? 'Admin' : 'Member'}
              </span>
            </button>
            {u.role === 'admin' ? <span style={{ fontSize: 11.5, color: C.faint }}>Included</span> : (
              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12.5, color: C.text, cursor: 'pointer' }}>
                <input type="checkbox" checked={!!u.can_build_docs} onChange={() => toggleBuildDocs(u)} />{u.can_build_docs ? 'Can build' : 'View only'}
              </label>
            )}
            {u.role === 'admin' ? <span style={{ fontSize: 11.5, color: C.faint }}>Included</span> : (
              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12.5, color: C.text, cursor: 'pointer' }}>
                <input type="checkbox" checked={!!u.can_build} onChange={() => toggleBuild(u)} />{u.can_build ? 'Can build' : 'View only'}
              </label>
            )}
            <button onClick={() => toggleActive(u)} disabled={u.id === currentUserId} title={u.active ? 'Deactivate' : 'Reactivate'} style={{
              display: 'flex', alignItems: 'center', gap: 5, border: `1px solid ${C.border}`, background: '#fff', borderRadius: 5,
              padding: '5px 10px', cursor: u.id === currentUserId ? 'default' : 'pointer', color: u.active ? C.red : C.green, fontSize: 12, fontFamily: 'inherit'
            }}>
              <Ban size={12} /> {u.active ? 'Deactivate' : 'Reactivate'}
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

// ================= Attachment tile =================
function AttachmentTile({ meta, url, isLoading, onLoad }) {
  const isImage = (meta.type || '').startsWith('image/');
  useEffect(() => { if (isImage && !url) onLoad(); /* eslint-disable-next-line */ }, []);
  function download() {
    if (url) { const a = document.createElement('a'); a.href = url; a.download = meta.name; a.click(); }
    else onLoad();
  }
  if (isImage) {
    return (
      <button onClick={download} style={{ border: `1px solid ${C.border}`, borderRadius: 6, overflow: 'hidden', background: '#fff', cursor: 'pointer', padding: 0, display: 'block', textAlign: 'left' }} title={`${meta.name} — click to download`}>
        {url ? <img src={url} alt={meta.name} style={{ width: '100%', height: 90, objectFit: 'cover', display: 'block' }} />
          : <div style={{ width: '100%', height: 90, display: 'flex', alignItems: 'center', justifyContent: 'center', background: C.bg }}>{isLoading ? <Loader2 size={16} color={C.faint} /> : <ImageIcon size={18} color={C.faint} />}</div>}
        <div style={{ padding: '6px 8px', fontSize: 11, color: C.dim, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{meta.name}</div>
      </button>
    );
  }
  return (
    <button onClick={download} style={{ border: `1px solid ${C.border}`, borderRadius: 6, background: '#fff', cursor: 'pointer', padding: '12px 10px', display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 6, textAlign: 'left' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, color: C.navy }}>{isLoading ? <Loader2 size={15} /> : <FileIcon size={15} />}<span style={{ fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 100 }}>{meta.name}</span></div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 10.5, color: C.faint }}><Download size={11} /> {fmtBytes(meta.size)}</div>
    </button>
  );
}

// Read-only display of one field's answer — used in the submission detail view for
// both flat (non-repeating) values and per-row values inside a repeating section.
function FieldAnswer({ field, value, attachmentUrlCache, loadingAttachments, onLoadAttachment }) {
  return (
    <div>
      <div style={{ fontSize: 11, fontWeight: 600, color: C.dim, textTransform: 'uppercase', letterSpacing: '0.03em', marginBottom: 4 }}>{field.label}</div>
      {field.type === 'Pass/Fail' ? (
        <PassFailBadge value={value} />
      ) : field.type === 'Photo' ? (
        value?.path ? (
          <div style={{ maxWidth: 150 }}>
            <AttachmentTile meta={value} url={attachmentUrlCache[value.id]} isLoading={!!loadingAttachments[value.id]} onLoad={() => onLoadAttachment(value)} />
          </div>
        ) : <span style={{ color: C.faint, fontSize: 13.5 }}>—</span>
      ) : (
        <div style={{ fontSize: 13.5, color: C.text, whiteSpace: 'pre-wrap' }}>{value || <span style={{ color: C.faint }}>—</span>}</div>
      )}
    </div>
  );
}

// Renders the correct input control for one field, given a flat value/onChange pair.
// Used both for a non-repeating section's fields and for one row inside a repeating one.
function FieldInput({ field, value, onChange, inputStyle }) {
  if (field.type === 'Text') return <input style={inputStyle} value={value || ''} onChange={e => onChange(e.target.value)} />;
  if (field.type === 'Notes') return <textarea style={{ ...inputStyle, minHeight: 80, resize: 'vertical' }} value={value || ''} onChange={e => onChange(e.target.value)} />;
  if (field.type === 'Number') return <input type="number" style={inputStyle} value={value || ''} onChange={e => onChange(e.target.value)} />;
  if (field.type === 'Date') return <input type="date" style={inputStyle} value={value || ''} onChange={e => onChange(e.target.value)} />;
  if (field.type === 'Yes/No') return <select style={inputStyle} value={value || ''} onChange={e => onChange(e.target.value)}><option value="">Select…</option><option>Yes</option><option>No</option></select>;
  if (field.type === 'Pass/Fail') return <select style={inputStyle} value={value || ''} onChange={e => onChange(e.target.value)}><option value="">Select…</option><option>Pass</option><option>Fail</option></select>;
  if (field.type === 'Multiple Choice') return (
    <select style={inputStyle} value={value || ''} onChange={e => onChange(e.target.value)}>
      <option value="">Select…</option>
      {(field.options || '').split(',').map(o => o.trim()).filter(Boolean).map(o => <option key={o}>{o}</option>)}
    </select>
  );
  if (field.type === 'Photo') return (
    <div>
      <input type="file" accept="image/*" capture="environment" onChange={e => onChange(e.target.files?.[0] || null)} style={inputStyle} />
      {value && <div style={{ marginTop: 6, display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: C.dim }}><Camera size={12} /> {value.name} ({fmtBytes(value.size)})</div>}
    </div>
  );
  return null;
}

// ================= Dashboard =================
function Dashboard({ kpis, statusChartData, typeChartData, expiringList, recentSubmissions, onSelectDoc, onSelectSubmission, total, onNew, canBuild }) {
  if (total === 0) return <EmptyState title="No documents yet" body={canBuild ? "Build your first SOP, checklist, or work instruction to start tracking it here." : "No one has added a document yet. Check back soon."} onNew={onNew} canCreate={canBuild} />;
  return (
    <div>
      <PageHeader title="Dashboard" subtitle="Live status across documents, templates, and submitted work." />
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 14, marginBottom: 28 }}>
        <KpiCard label="Total Documents" value={kpis.total} color={C.navy} />
        <KpiCard label="Approved" value={kpis.approved} color={C.green} />
        <KpiCard label="Pending Approval" value={kpis.pending} color={C.amber} />
        <KpiCard label="Expiring ≤ 30 Days" value={kpis.expiringSoon} color={C.amber} />
        <KpiCard label="Expired" value={kpis.expired} color={C.red} />
        <KpiCard label="Submissions (7 days)" value={kpis.submissionsThisWeek} color={C.blue} />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1.1fr 1fr', gap: 18, marginBottom: 18 }}>
        <ChartCard title="Documents by Status">
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={statusChartData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={85} paddingAngle={2}>
                {statusChartData.map((entry) => <Cell key={entry.name} fill={statusColor(entry.name)} />)}
              </Pie>
              <Tooltip /><Legend verticalAlign="bottom" height={28} wrapperStyle={{ fontSize: 12 }} />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>
        <ChartCard title="Documents by Type">
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={typeChartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11, fill: C.dim }} interval={0} angle={-20} textAnchor="end" height={50} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: C.dim }} />
              <Tooltip /><Bar dataKey="count" fill={C.navy} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 18 }}>
        <ChartCard title="Expiring Soon">
          {expiringList.length === 0 ? <div style={{ color: C.faint, fontSize: 13.5, padding: '8px 2px' }}>Nothing expiring in the next 30 days.</div> : (
            <div>
              {expiringList.map(d => (
                <button key={d.id} onClick={() => onSelectDoc(d)} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', padding: '10px 4px', border: 'none', borderTop: `1px solid ${C.border}`, background: 'transparent', cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                    <span style={{ fontSize: 13.5, fontWeight: 500, color: C.text }}>{d.title}</span>
                    <span style={{ fontSize: 11.5, color: C.faint }}>{d.type} · {d.department || 'No department'}</span>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <span style={{ fontSize: 12, fontWeight: 600, color: d.days <= 7 ? C.red : C.amber }}>{d.days === 0 ? 'Expires today' : `${d.days} day${d.days === 1 ? '' : 's'} left`}</span>
                    <ChevronRight size={14} color={C.faint} />
                  </div>
                </button>
              ))}
            </div>
          )}
        </ChartCard>
        <ChartCard title="Recent Submissions">
          {recentSubmissions.length === 0 ? <div style={{ color: C.faint, fontSize: 13.5, padding: '8px 2px' }}>No forms filled out yet.</div> : (
            <div>
              {recentSubmissions.map(s => (
                <button key={s.id} onClick={() => onSelectSubmission(s)} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', padding: '10px 4px', border: 'none', borderTop: `1px solid ${C.border}`, background: 'transparent', cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                    <span style={{ fontSize: 13.5, fontWeight: 500, color: C.text }}>{s.title || s.templateName}</span>
                    <span style={{ fontSize: 11.5, color: C.faint }}>By {s.filledBy} · {s.filledAt}</span>
                  </div>
                  <ChevronRight size={14} color={C.faint} />
                </button>
              ))}
            </div>
          )}
        </ChartCard>
      </div>
    </div>
  );
}
function statusColor(name) {
  if (name === 'Approved') return C.green;
  if (name === 'Pending Approval') return C.amber;
  if (name === 'Expired') return C.red;
  return C.faint;
}

// ================= Documents List =================
function DocumentsList({ filtered, search, setSearch, filterType, setFilterType, filterStatus, setFilterStatus, onEdit, onDelete, onView, onQuickStatus, hasAny, onNew, isAdmin, canBuild }) {
  if (!hasAny) return <EmptyState title="No documents yet" body={canBuild ? "Build your first SOP, checklist, or work instruction to start tracking it here." : "No one has added a document yet. Check back soon."} onNew={onNew} canCreate={canBuild} />;
  return (
    <div>
      <PageHeader title="Documents" subtitle={`${filtered.length} document${filtered.length === 1 ? '' : 's'}`} />
      <div style={{ display: 'flex', gap: 10, marginBottom: 18, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative', flex: '1 1 220px' }}>
          <Search size={14} color={C.faint} style={{ position: 'absolute', left: 10, top: 11 }} />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search title, department, or owner" style={{ width: '100%', padding: '9px 12px 9px 32px', borderRadius: 5, border: `1px solid ${C.border}`, fontSize: 13.5, fontFamily: 'inherit' }} />
        </div>
        <select value={filterType} onChange={e => setFilterType(e.target.value)} style={selectStyle(C)}><option>All</option>{TYPE_OPTIONS.map(t => <option key={t}>{t}</option>)}</select>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={selectStyle(C)}><option>All</option>{[...STATUS_OPTIONS, 'Expired'].map(s => <option key={s}>{s}</option>)}</select>
      </div>
      <div style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'hidden' }}>
        <div style={{ display: 'grid', gridTemplateColumns: isAdmin ? '2.2fr 1fr 1fr 1fr 1fr 1.3fr 90px' : '2.2fr 1fr 1fr 1fr 1fr 1.3fr 46px', padding: '10px 16px', fontSize: 11, fontWeight: 600, color: C.faint, textTransform: 'uppercase', letterSpacing: '0.04em', borderBottom: `1px solid ${C.border}`, background: C.bg }}>
          <span>Title</span><span>Type</span><span>Department</span><span>Revision</span><span>Expires</span><span>Status</span><span></span>
        </div>
        {filtered.length === 0 ? <div style={{ padding: '32px 16px', textAlign: 'center', color: C.faint, fontSize: 13.5 }}>No documents match these filters.</div> : filtered.map(d => (
          <div key={d.id} style={{ display: 'grid', gridTemplateColumns: isAdmin ? '2.2fr 1fr 1fr 1fr 1fr 1.3fr 90px' : '2.2fr 1fr 1fr 1fr 1fr 1.3fr 46px', padding: '13px 16px', fontSize: 13.5, borderBottom: `1px solid ${C.border}`, alignItems: 'center' }}>
            <button onClick={() => onView(d)} style={{ border: 'none', background: 'none', textAlign: 'left', padding: 0, cursor: 'pointer', fontWeight: 600, color: C.navy, fontFamily: 'inherit', fontSize: 13.5, display: 'flex', alignItems: 'center', gap: 6 }}>
              {d.title}{d.attachments && d.attachments.length > 0 && (<span style={{ display: 'inline-flex', alignItems: 'center', gap: 2, color: C.faint, fontWeight: 500, fontSize: 11.5 }}><Paperclip size={11} />{d.attachments.length}</span>)}
            </button>
            <span style={{ color: C.dim }}>{d.type}</span>
            <span style={{ color: C.dim }}>{d.department || '—'}</span>
            <span style={{ color: C.dim, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{d.revision || 'A'}</span>
            <span style={{ color: C.dim, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{d.expiryDate || '—'}</span>
            {canBuild ? (
              <select value={d.effStatus === 'Expired' ? 'Approved' : d.effStatus} onChange={e => onQuickStatus(d.id, e.target.value)} style={{ border: 'none', background: 'none', cursor: 'pointer', fontFamily: 'inherit' }}>{STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}</select>
            ) : (<Badge status={d.effStatus} />)}
            <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
              {canBuild && <IconBtn onClick={() => onEdit(d)} title="Edit"><Pencil size={14} /></IconBtn>}
              {isAdmin && <IconBtn onClick={() => onDelete(d.id)} title="Delete"><Trash2 size={14} /></IconBtn>}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ================= Document Form =================
function DocumentForm({ form, setForm, onCancel, onSubmit, onRemoveExisting, inputStyle, labelStyle, saving }) {
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [fileError, setFileError] = useState(null);
  const [wordError, setWordError] = useState(null);
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const wordInputRef = useRef(null);
  function handleFiles(e) {
    const files = Array.from(e.target.files || []);
    setFileError(null);
    files.forEach(file => {
      if (file.size > MAX_FILE_BYTES) { setFileError(`"${file.name}" is too large (limit ${fmtBytes(MAX_FILE_BYTES)}).`); return; }
      setForm(f => ({ ...f, attachments: [...(f.attachments || []), { id: uid('a'), name: file.name, type: file.type, size: file.size, file, isNew: true }] }));
    });
    if (fileInputRef.current) fileInputRef.current.value = '';
  }
  function removeAttachment(att) {
    setForm(f => ({ ...f, attachments: (f.attachments || []).filter(a => a.id !== att.id) }));
    if (!att.isNew && att.path) onRemoveExisting(att.path);
  }
  async function handleWordFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setWordError(null); setImporting(true);
    try {
      const text = await extractWordText(file);
      setForm(f => ({ ...f, title: f.title || titleFromFilename(file.name), content: text }));
    } catch (err) { setWordError('Could not read that file. Make sure it\'s a .docx Word document.'); }
    finally { setImporting(false); if (wordInputRef.current) wordInputRef.current.value = ''; }
  }
  return (
    <div style={{ maxWidth: 640 }}>
      <PageHeader title={form.id ? 'Edit Document' : 'New Document'} subtitle="Fields marked with an asterisk are required." />
      <div style={{ marginBottom: 16, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <button type="button" onClick={() => wordInputRef.current && wordInputRef.current.click()} disabled={importing} style={{ ...btnGhost(C), padding: '8px 14px', fontSize: 13 }}>
          {importing ? <Loader2 size={14} /> : <Upload size={14} />} {importing ? 'Reading file…' : 'Import from Word (.docx)'}
        </button>
        <input ref={wordInputRef} type="file" accept=".docx" onChange={handleWordFile} style={{ display: 'none' }} />
        {wordError && <span style={{ color: C.red, fontSize: 12 }}>{wordError}</span>}
      </div>
      <form onSubmit={onSubmit} style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, padding: 24 }}>
        <div style={{ marginBottom: 16 }}><label style={labelStyle}>Title *</label>
          <input required style={inputStyle} value={form.title} onChange={e => set('title', e.target.value)} placeholder="e.g. Weld Inspection Checklist — Line 3" /></div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
          <div><label style={labelStyle}>Document Type</label><select style={inputStyle} value={form.type} onChange={e => set('type', e.target.value)}>{TYPE_OPTIONS.map(t => <option key={t}>{t}</option>)}</select></div>
          <div><label style={labelStyle}>Status</label><select style={inputStyle} value={form.status} onChange={e => set('status', e.target.value)}>{STATUS_OPTIONS.map(s => <option key={s}>{s}</option>)}</select></div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
          <div><label style={labelStyle}>Department</label><input style={inputStyle} value={form.department} onChange={e => set('department', e.target.value)} placeholder="e.g. Assembly" /></div>
          <div><label style={labelStyle}>Owner</label><input style={inputStyle} value={form.owner} onChange={e => set('owner', e.target.value)} placeholder="e.g. F. Felix" /></div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 16, marginBottom: 16 }}>
          <div><label style={labelStyle}>Revision</label><input style={inputStyle} value={form.revision} onChange={e => set('revision', e.target.value)} placeholder="A" /></div>
          <div><label style={labelStyle}>Effective Date</label><input type="date" style={inputStyle} value={form.effectiveDate} onChange={e => set('effectiveDate', e.target.value)} /></div>
          <div><label style={labelStyle}>Expiry Date</label><input type="date" style={inputStyle} value={form.expiryDate} onChange={e => set('expiryDate', e.target.value)} /></div>
        </div>
        <div style={{ marginBottom: 20 }}><label style={labelStyle}>Content</label>
          <textarea style={{ ...inputStyle, minHeight: 160, resize: 'vertical', fontFamily: 'inherit', lineHeight: 1.6 }} value={form.content} onChange={e => set('content', e.target.value)} placeholder="Steps, checklist items, or the body of the document..." /></div>
        <div style={{ marginBottom: 22 }}>
          <label style={labelStyle}>Attachments</label>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginBottom: (form.attachments || []).length > 0 ? 12 : 0 }}>
            {(form.attachments || []).map(a => (
              <div key={a.id} style={{ display: 'flex', alignItems: 'center', gap: 8, border: `1px solid ${C.border}`, borderRadius: 6, padding: '7px 10px', fontSize: 12.5, background: C.bg }}>
                {(a.type || '').startsWith('image/') ? <ImageIcon size={14} color={C.dim} /> : <FileIcon size={14} color={C.dim} />}
                <span style={{ maxWidth: 140, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{a.name}</span>
                <span style={{ color: C.faint }}>{fmtBytes(a.size)}</span>
                <button type="button" onClick={() => removeAttachment(a)} style={{ border: 'none', background: 'none', cursor: 'pointer', color: C.faint, display: 'flex' }}><X size={13} /></button>
              </div>
            ))}
          </div>
          <button type="button" onClick={() => fileInputRef.current && fileInputRef.current.click()} style={{ ...btnGhost(C), padding: '8px 14px', fontSize: 13 }}><Paperclip size={14} /> Add Photo or File</button>
          <input ref={fileInputRef} type="file" multiple accept="image/*,.pdf,.doc,.docx,.xls,.xlsx" onChange={handleFiles} style={{ display: 'none' }} />
          {fileError && <div style={{ color: C.red, fontSize: 12, marginTop: 8 }}>{fileError}</div>}
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button type="submit" disabled={saving} style={{ ...btnPrimary(C), opacity: saving ? 0.7 : 1 }}>{saving ? <Loader2 size={14} /> : null}{saving ? 'Saving…' : (form.id ? 'Save Changes' : 'Create Document')}</button>
          <button type="button" onClick={onCancel} style={btnGhost(C)}>Cancel</button>
        </div>
      </form>
    </div>
  );
}

// ================= Templates list =================
function TemplatesList({ templates, isAdmin, canBuild, onNew, onEdit, onDelete, onFill, submissions }) {
  function countFor(id) { return submissions.filter(s => s.templateId === id).length; }
  if (templates.length === 0) {
    return (
      <div>
        <PageHeader title="Templates" subtitle="Reusable forms your team fills out over and over." />
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', background: '#fff', border: `1px dashed ${C.borderStrong}`, borderRadius: 10, padding: '48px 40px', maxWidth: 480 }}>
          <ClipboardList size={28} color={C.faint} />
          <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 19, margin: '16px 0 6px' }}>No templates yet</h2>
          <p style={{ fontSize: 13.5, color: C.dim, margin: '0 0 18px', lineHeight: 1.6 }}>
            {canBuild ? 'Build a checklist or inspection form once, and anyone can fill it out from here going forward.' : 'No one has built a form yet. Check back soon.'}
          </p>
          {canBuild && <button onClick={onNew} style={btnPrimary(C)}><PlusCircle size={14} /> New Template</button>}
        </div>
      </div>
    );
  }
  return (
    <div>
      <PageHeader title="Templates" subtitle="Reusable forms your team fills out over and over." />
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 14 }}>
        {templates.map(t => {
          const fieldCount = (t.sections || []).reduce((sum, s) => sum + s.fields.length, 0);
          return (
            <div key={t.id} style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, padding: 20 }}>
              <div style={{ fontSize: 11, color: C.faint, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 6 }}>{t.category}{t.department ? ` · ${t.department}` : ''}</div>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 16, fontWeight: 600, marginBottom: 8 }}>{t.name}</div>
              <div style={{ fontSize: 12.5, color: C.dim, marginBottom: 16 }}>
                {(t.sections || []).length} section{(t.sections || []).length === 1 ? '' : 's'} · {fieldCount} field{fieldCount === 1 ? '' : 's'} · filled out {countFor(t.id)} time{countFor(t.id) === 1 ? '' : 's'}
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={() => onFill(t)} style={{ ...btnPrimary(C), padding: '8px 14px', fontSize: 13 }}><ClipboardCheck size={14} /> Fill Out</button>
                {canBuild && <IconBtn onClick={() => onEdit(t)} title="Edit"><Pencil size={14} /></IconBtn>}
                {isAdmin && <IconBtn onClick={() => onDelete(t.id)} title="Delete"><Trash2 size={14} /></IconBtn>}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ================= Template builder =================
function TemplateForm({ templateForm, setTemplateForm, onCancel, onSubmit, onRemoveExisting, inputStyle, labelStyle, saving }) {
  const set = (k, v) => setTemplateForm(f => ({ ...f, [k]: v }));
  const [fileError, setFileError] = useState(null);
  const [wordError, setWordError] = useState(null);
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);
  const wordInputRef = useRef(null);

  function addSection() { setTemplateForm(f => ({ ...f, sections: [...f.sections, { id: uid('sec'), title: '', repeating: false, fields: [] }] })); }
  function updateSection(id, patch) { setTemplateForm(f => ({ ...f, sections: f.sections.map(s => s.id === id ? { ...s, ...patch } : s) })); }
  function removeSection(id) { setTemplateForm(f => ({ ...f, sections: f.sections.filter(s => s.id !== id) })); }
  function moveSection(id, dir) {
    setTemplateForm(f => {
      const idx = f.sections.findIndex(s => s.id === id);
      const swap = idx + dir;
      if (swap < 0 || swap >= f.sections.length) return f;
      const next = [...f.sections];
      [next[idx], next[swap]] = [next[swap], next[idx]];
      return { ...f, sections: next };
    });
  }
  function addFieldToSection(sectionId) {
    setTemplateForm(f => ({ ...f, sections: f.sections.map(s => s.id === sectionId ? { ...s, fields: [...s.fields, { id: uid('f'), label: '', type: 'Text', options: '', required: true, condition: null }] } : s) }));
  }
  function updateFieldInSection(sectionId, fieldId, patch) {
    setTemplateForm(f => ({ ...f, sections: f.sections.map(s => s.id !== sectionId ? s : { ...s, fields: s.fields.map(fl => fl.id === fieldId ? { ...fl, ...patch } : fl) }) }));
  }
  function removeFieldFromSection(sectionId, fieldId) {
    setTemplateForm(f => ({
      ...f,
      sections: f.sections.map(s => s.id !== sectionId ? s : {
        ...s,
        fields: s.fields.filter(fl => fl.id !== fieldId).map(fl => fl.condition && fl.condition.fieldId === fieldId ? { ...fl, condition: null } : fl),
      }),
    }));
  }
  function moveFieldInSection(sectionId, fieldId, dir) {
    setTemplateForm(f => ({
      ...f,
      sections: f.sections.map(s => {
        if (s.id !== sectionId) return s;
        const idx = s.fields.findIndex(fl => fl.id === fieldId);
        const swap = idx + dir;
        if (swap < 0 || swap >= s.fields.length) return s;
        const next = [...s.fields];
        [next[idx], next[swap]] = [next[swap], next[idx]];
        return { ...s, fields: next };
      }),
    }));
  }
  function handleFiles(e) {
    const files = Array.from(e.target.files || []);
    setFileError(null);
    files.forEach(file => {
      if (file.size > MAX_FILE_BYTES) { setFileError(`"${file.name}" is too large (limit ${fmtBytes(MAX_FILE_BYTES)}).`); return; }
      setTemplateForm(f => ({ ...f, attachments: [...(f.attachments || []), { id: uid('a'), name: file.name, type: file.type, size: file.size, file, isNew: true }] }));
    });
    if (fileInputRef.current) fileInputRef.current.value = '';
  }
  function removeAttachment(att) {
    setTemplateForm(f => ({ ...f, attachments: (f.attachments || []).filter(a => a.id !== att.id) }));
    if (!att.isNew && att.path) onRemoveExisting(att.path);
  }
  async function handleWordFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setWordError(null); setImporting(true);
    try {
      const text = await extractWordText(file);
      setTemplateForm(f => ({ ...f, name: f.name || titleFromFilename(file.name), description: text }));
    } catch (err) { setWordError('Could not read that file. Make sure it\'s a .docx Word document.'); }
    finally { setImporting(false); if (wordInputRef.current) wordInputRef.current.value = ''; }
  }

  const canSubmit = templateForm.name.trim() && templateForm.sections.length > 0 &&
    templateForm.sections.every(s => s.title.trim() && s.fields.length > 0 && s.fields.every(f => f.label.trim()));

  return (
    <div style={{ maxWidth: 700 }}>
      <PageHeader title={templateForm.id ? 'Edit Template' : 'New Template'} subtitle="Group related fields into sections — mark a section \u201crepeating\u201d if people should be able to log more than one of it." />
      <div style={{ marginBottom: 16, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <button type="button" onClick={() => wordInputRef.current && wordInputRef.current.click()} disabled={importing} style={{ ...btnGhost(C), padding: '8px 14px', fontSize: 13 }}>
          {importing ? <Loader2 size={14} /> : <Upload size={14} />} {importing ? 'Reading file…' : 'Import from Word (.docx)'}
        </button>
        <input ref={wordInputRef} type="file" accept=".docx" onChange={handleWordFile} style={{ display: 'none' }} />
        {wordError && <span style={{ color: C.red, fontSize: 12 }}>{wordError}</span>}
        <span style={{ fontSize: 11.5, color: C.faint }}>Fills in the name and description below — add your sections and fields manually after.</span>
      </div>
      <form onSubmit={onSubmit} style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, padding: 24 }}>
        <div style={{ marginBottom: 16 }}>
          <label style={labelStyle}>Template Name *</label>
          <input required style={inputStyle} value={templateForm.name} onChange={e => set('name', e.target.value)} placeholder="e.g. Building Inspection Form" />
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
          <div><label style={labelStyle}>Category</label>
            <select style={inputStyle} value={templateForm.category} onChange={e => set('category', e.target.value)}>{TYPE_OPTIONS.map(t => <option key={t}>{t}</option>)}</select></div>
          <div><label style={labelStyle}>Department</label>
            <input style={inputStyle} value={templateForm.department} onChange={e => set('department', e.target.value)} placeholder="e.g. Assembly" /></div>
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={labelStyle}>Submission Title Template</label>
          <input style={inputStyle} value={templateForm.titleTemplate || ''} onChange={e => set('titleTemplate', e.target.value)} placeholder="e.g. {{Building Number}} - {{Location}}" />
          <div style={{ fontSize: 11.5, color: C.faint, marginTop: 5 }}>Optional. Wrap a field's exact label in double braces to pull its answer into each submission's title — makes entries easier to tell apart in the Submissions list.</div>
        </div>
        <div style={{ marginBottom: 22 }}>
          <label style={labelStyle}>Description / Instructions</label>
          <textarea style={{ ...inputStyle, minHeight: 90, resize: 'vertical', fontFamily: 'inherit', lineHeight: 1.6 }} value={templateForm.description || ''} onChange={e => set('description', e.target.value)} placeholder="Context or instructions shown before the fields when someone fills this out..." />
        </div>

        <div style={{ marginBottom: 22 }}>
          <label style={labelStyle}>Reference Attachments</label>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginBottom: (templateForm.attachments || []).length > 0 ? 12 : 0 }}>
            {(templateForm.attachments || []).map(a => (
              <div key={a.id} style={{ display: 'flex', alignItems: 'center', gap: 8, border: `1px solid ${C.border}`, borderRadius: 6, padding: '7px 10px', fontSize: 12.5, background: C.bg }}>
                {(a.type || '').startsWith('image/') ? <ImageIcon size={14} color={C.dim} /> : <FileIcon size={14} color={C.dim} />}
                <span style={{ maxWidth: 140, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{a.name}</span>
                <span style={{ color: C.faint }}>{fmtBytes(a.size)}</span>
                <button type="button" onClick={() => removeAttachment(a)} style={{ border: 'none', background: 'none', cursor: 'pointer', color: C.faint, display: 'flex' }}><X size={13} /></button>
              </div>
            ))}
          </div>
          <button type="button" onClick={() => fileInputRef.current && fileInputRef.current.click()} style={{ ...btnGhost(C), padding: '8px 14px', fontSize: 13 }}><Paperclip size={14} /> Add Photo or File</button>
          <input ref={fileInputRef} type="file" multiple accept="image/*,.pdf,.doc,.docx,.xls,.xlsx" onChange={handleFiles} style={{ display: 'none' }} />
          {fileError && <div style={{ color: C.red, fontSize: 12, marginTop: 8 }}>{fileError}</div>}
        </div>

        <label style={labelStyle}>Sections</label>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16, marginBottom: 14 }}>
          {templateForm.sections.length === 0 && <div style={{ fontSize: 13, color: C.faint, padding: '12px 2px' }}>No sections yet — add one below to start adding fields.</div>}
          {templateForm.sections.map((section, si) => (
            <div key={section.id} style={{ border: `1px solid ${C.borderStrong}`, borderRadius: 8, padding: 16, background: '#fbfcfd' }}>
              <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap', marginBottom: 10 }}>
                <Layers size={15} color={C.navy} style={{ flexShrink: 0 }} />
                <input
                  style={{ ...inputStyle, flex: '2 1 180px', background: '#fff', fontWeight: 600 }}
                  value={section.title} onChange={e => updateSection(section.id, { title: e.target.value })}
                  placeholder="Section title, e.g. Building Details"
                />
                <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12.5, color: C.dim, whiteSpace: 'nowrap' }}>
                  <input type="checkbox" checked={!!section.repeating} onChange={e => updateSection(section.id, { repeating: e.target.checked })} />
                  Repeating (multiple entries)
                </label>
                <div style={{ display: 'flex', gap: 4, marginLeft: 'auto' }}>
                  <IconBtn onClick={() => moveSection(section.id, -1)} title="Move section up"><ArrowUp size={13} /></IconBtn>
                  <IconBtn onClick={() => moveSection(section.id, 1)} title="Move section down"><ArrowDown size={13} /></IconBtn>
                  <IconBtn onClick={() => removeSection(section.id)} title="Remove section"><Trash2 size={13} /></IconBtn>
                </div>
              </div>
              {section.repeating && (
                <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 12, display: 'flex', alignItems: 'center', gap: 5 }}>
                  <Copy size={11} /> People filling this out can add as many entries of this section as they need — great for logging a list of defects, items, or issues.
                </div>
              )}

              <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginBottom: 10 }}>
                {section.fields.length === 0 && <div style={{ fontSize: 12.5, color: C.faint, padding: '4px 2px' }}>No fields in this section yet.</div>}
                {section.fields.map((f, fi) => {
                  const eligible = section.fields.slice(0, fi).filter(other => CONDITIONABLE_TYPES.includes(other.type) && other.label.trim());
                  const conditionField = f.condition ? section.fields.find(x => x.id === f.condition.fieldId) : null;
                  return (
                    <div key={f.id} style={{ border: `1px solid ${C.border}`, borderRadius: 6, padding: 12, background: C.bg }}>
                      <div style={{ display: 'flex', gap: 10, marginBottom: f.type === 'Multiple Choice' ? 10 : 0, flexWrap: 'wrap' }}>
                        <input style={{ ...inputStyle, flex: '2 1 160px', background: '#fff' }} value={f.label} onChange={e => updateFieldInSection(section.id, f.id, { label: e.target.value })} placeholder="Field label, e.g. Torque spec met?" />
                        <select style={{ ...inputStyle, flex: '1 1 130px', background: '#fff' }} value={f.type} onChange={e => updateFieldInSection(section.id, f.id, { type: e.target.value })}>{FIELD_TYPES.map(ft => <option key={ft}>{ft}</option>)}</select>
                        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12.5, color: C.dim, whiteSpace: 'nowrap' }}>
                          <input type="checkbox" checked={f.required} onChange={e => updateFieldInSection(section.id, f.id, { required: e.target.checked })} /> Required
                        </label>
                        <div style={{ display: 'flex', gap: 4 }}>
                          <IconBtn onClick={() => moveFieldInSection(section.id, f.id, -1)} title="Move up"><ArrowUp size={13} /></IconBtn>
                          <IconBtn onClick={() => moveFieldInSection(section.id, f.id, 1)} title="Move down"><ArrowDown size={13} /></IconBtn>
                          <IconBtn onClick={() => removeFieldFromSection(section.id, f.id)} title="Remove field"><X size={14} /></IconBtn>
                        </div>
                      </div>
                      {f.type === 'Multiple Choice' && (
                        <input style={{ ...inputStyle, background: '#fff', marginBottom: 10 }} value={f.options} onChange={e => updateFieldInSection(section.id, f.id, { options: e.target.value })} placeholder="Options, separated by commas (e.g. Line 1, Line 2, Line 3)" />
                      )}
                      {f.type === 'Photo' && (
                        <div style={{ fontSize: 12, color: C.faint, display: 'flex', alignItems: 'center', gap: 6, marginBottom: 10 }}><Camera size={13} /> Whoever fills this out will be asked to attach a photo here.</div>
                      )}
                      <div style={{ paddingTop: 10, borderTop: `1px dashed ${C.borderStrong}` }}>
                        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: C.dim, cursor: eligible.length === 0 ? 'default' : 'pointer' }}>
                          <input type="checkbox" checked={!!f.condition} disabled={eligible.length === 0} onChange={e => updateFieldInSection(section.id, f.id, { condition: e.target.checked ? { fieldId: '', equals: '' } : null })} />
                          Only show this field conditionally
                        </label>
                        {eligible.length === 0 && <div style={{ fontSize: 11, color: C.faint, marginTop: 4 }}>Add a Yes/No, Pass/Fail, or Multiple Choice field above this one (in the same section) to make it conditional.</div>}
                        {f.condition && (
                          <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                            <span style={{ fontSize: 12, color: C.dim }}>Show when</span>
                            <select style={{ ...inputStyle, background: '#fff', flex: '1 1 150px', padding: '7px 10px' }} value={f.condition.fieldId} onChange={e => updateFieldInSection(section.id, f.id, { condition: { fieldId: e.target.value, equals: '' } })}>
                              <option value="">Choose a field…</option>
                              {eligible.map(ef => <option key={ef.id} value={ef.id}>{ef.label}</option>)}
                            </select>
                            <span style={{ fontSize: 12, color: C.dim }}>is</span>
                            <select style={{ ...inputStyle, background: '#fff', flex: '1 1 130px', padding: '7px 10px' }} value={f.condition.equals} disabled={!f.condition.fieldId} onChange={e => updateFieldInSection(section.id, f.id, { condition: { ...f.condition, equals: e.target.value } })}>
                              <option value="">Select…</option>
                              {optionsForField(conditionField).map(o => <option key={o}>{o}</option>)}
                            </select>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
              <button type="button" onClick={() => addFieldToSection(section.id)} style={{ ...btnGhost(C), padding: '7px 12px', fontSize: 12.5 }}><PlusCircle size={13} /> Add Field</button>
            </div>
          ))}
        </div>
        <button type="button" onClick={addSection} style={{ ...btnGhost(C), padding: '8px 14px', fontSize: 13, marginBottom: 22 }}><FolderPlus size={14} /> Add Section</button>

        <div style={{ display: 'flex', gap: 10 }}>
          <button type="submit" disabled={!canSubmit || saving} style={{ ...btnPrimary(C), opacity: (!canSubmit || saving) ? 0.6 : 1 }}>
            {saving ? <Loader2 size={14} /> : null}{saving ? 'Saving…' : (templateForm.id ? 'Save Template' : 'Create Template')}
          </button>
          <button type="button" onClick={onCancel} style={btnGhost(C)}>Cancel</button>
        </div>
      </form>
    </div>
  );
}

// ================= Fill out a template =================
function FillForm({ template, onCancel, onSubmit, inputStyle, labelStyle, saving, sessionName }) {
  const [values, setValues] = useState(() => {
    const init = {};
    (template.sections || []).forEach(s => { if (s.repeating) init[s.id] = []; });
    return init;
  });
  const [errors, setErrors] = useState([]);
  const [refCache, setRefCache] = useState({});
  const [refLoading, setRefLoading] = useState({});

  function setVal(fieldId, v) { setValues(prev => ({ ...prev, [fieldId]: v })); }
  function setRowVal(sectionId, rowIdx, fieldId, v) {
    setValues(prev => {
      const rows = [...(prev[sectionId] || [])];
      rows[rowIdx] = { ...rows[rowIdx], [fieldId]: v };
      return { ...prev, [sectionId]: rows };
    });
  }
  function addRow(sectionId) { setValues(prev => ({ ...prev, [sectionId]: [...(prev[sectionId] || []), {}] })); }
  function removeRow(sectionId, rowIdx) { setValues(prev => ({ ...prev, [sectionId]: (prev[sectionId] || []).filter((_, i) => i !== rowIdx) })); }

  async function loadRef(att) {
    if (refCache[att.id] || refLoading[att.id] || !att.path) return;
    setRefLoading(s => ({ ...s, [att.id]: true }));
    try {
      const { data, error } = await supabase.storage.from(ATTACH_BUCKET).download(att.path);
      if (!error && data) setRefCache(c => ({ ...c, [att.id]: URL.createObjectURL(data) }));
    } finally { setRefLoading(s => ({ ...s, [att.id]: false })); }
  }

  function handleSubmit(e) {
    e.preventDefault();
    const missing = [];
    template.sections.forEach(section => {
      if (section.repeating) {
        (values[section.id] || []).forEach((row, i) => {
          section.fields.filter(f => fieldVisible(f, row)).forEach(f => {
            if (f.required && (f.type === 'Photo' ? !row[f.id] : !(row[f.id] || '').toString().trim())) {
              missing.push(`${f.label} (${section.title || 'entry'} #${i + 1})`);
            }
          });
        });
      } else {
        section.fields.filter(f => fieldVisible(f, values)).forEach(f => {
          if (f.required && (f.type === 'Photo' ? !values[f.id] : !(values[f.id] || '').toString().trim())) {
            missing.push(f.label);
          }
        });
      }
    });
    if (missing.length > 0) { setErrors(missing); return; }
    setErrors([]);
    onSubmit(values);
  }

  return (
    <div style={{ maxWidth: 680 }}>
      <PageHeader title={template.name} subtitle={`${template.category}${template.department ? ' · ' + template.department : ''} — filling out as ${sessionName}`} />

      {(template.description || (template.attachments && template.attachments.length > 0)) && (
        <div style={{ background: C.blueBg, border: `1px solid ${C.border}`, borderRadius: 8, padding: 18, marginBottom: 18 }}>
          {template.description && <div style={{ fontSize: 13.5, color: C.text, whiteSpace: 'pre-wrap', lineHeight: 1.6, marginBottom: (template.attachments || []).length > 0 ? 14 : 0 }}>{template.description}</div>}
          {(template.attachments && template.attachments.length > 0) && (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(120px, 1fr))', gap: 10 }}>
              {template.attachments.map(a => <AttachmentTile key={a.id} meta={a} url={refCache[a.id]} isLoading={!!refLoading[a.id]} onLoad={() => loadRef(a)} />)}
            </div>
          )}
        </div>
      )}

      <form onSubmit={handleSubmit} style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, padding: 24 }}>
        {errors.length > 0 && (
          <div style={{ background: C.redBg, color: C.red, borderRadius: 6, padding: '10px 14px', fontSize: 13, marginBottom: 18 }}>
            Fill in required field{errors.length === 1 ? '' : 's'}: {errors.join(', ')}
          </div>
        )}

        <div style={{ display: 'flex', flexDirection: 'column', gap: 26 }}>
          {template.sections.map(section => (
            <div key={section.id}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.navy, marginBottom: 12, paddingBottom: 8, borderBottom: `2px solid ${C.border}` }}>
                {section.title || 'Section'}
              </div>
              {section.repeating ? (
                <div>
                  {(values[section.id] || []).length === 0 && (
                    <div style={{ fontSize: 12.5, color: C.faint, marginBottom: 12 }}>No entries yet. Add one if this section applies.</div>
                  )}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 12 }}>
                    {(values[section.id] || []).map((row, rowIdx) => (
                      <div key={rowIdx} style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 16, background: C.bg }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                          <span style={{ fontSize: 11.5, fontWeight: 700, color: C.dim, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Entry {rowIdx + 1}</span>
                          <button type="button" onClick={() => removeRow(section.id, rowIdx)} style={{ border: 'none', background: 'none', color: C.red, fontSize: 12, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
                            <Trash2 size={12} /> Remove
                          </button>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                          {section.fields.filter(f => fieldVisible(f, row)).map(f => (
                            <div key={f.id}>
                              <label style={labelStyle}>{f.label}{f.required && ' *'}</label>
                              <FieldInput field={f} value={row[f.id]} onChange={v => setRowVal(section.id, rowIdx, f.id, v)} inputStyle={inputStyle} />
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                  <button type="button" onClick={() => addRow(section.id)} style={{ ...btnGhost(C), padding: '8px 14px', fontSize: 13 }}>
                    <PlusCircle size={14} /> Add {section.title || 'Entry'}
                  </button>
                </div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
                  {section.fields.filter(f => fieldVisible(f, values)).map(f => (
                    <div key={f.id}>
                      <label style={labelStyle}>{f.label}{f.required && ' *'}</label>
                      <FieldInput field={f} value={values[f.id]} onChange={v => setVal(f.id, v)} inputStyle={inputStyle} />
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>

        <div style={{ display: 'flex', gap: 10, marginTop: 26 }}>
          <button type="submit" disabled={saving} style={{ ...btnPrimary(C), opacity: saving ? 0.7 : 1 }}>{saving ? <Loader2 size={14} /> : null}{saving ? 'Submitting…' : 'Submit'}</button>
          <button type="button" onClick={onCancel} style={btnGhost(C)}>Cancel</button>
        </div>
      </form>
    </div>
  );
}

// ================= Submissions list =================
function SubmissionsList({ submissions, templates, onView }) {
  const [filterTemplate, setFilterTemplate] = useState('All');
  const filtered = filterTemplate === 'All' ? submissions : submissions.filter(s => s.templateId === filterTemplate);
  const sorted = [...filtered].sort((a, b) => new Date(b.filledAtTime || b.filledAt) - new Date(a.filledAtTime || a.filledAt));

  if (submissions.length === 0) {
    return (
      <div>
        <PageHeader title="Submissions" subtitle="Every completed template, saved as its own record." />
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', background: '#fff', border: `1px dashed ${C.borderStrong}`, borderRadius: 10, padding: '48px 40px', maxWidth: 480 }}>
          <ListChecks size={28} color={C.faint} />
          <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 19, margin: '16px 0 6px' }}>No submissions yet</h2>
          <p style={{ fontSize: 13.5, color: C.dim, margin: 0, lineHeight: 1.6 }}>Once someone fills out a template from the Templates tab, it'll show up here.</p>
        </div>
      </div>
    );
  }
  return (
    <div>
      <PageHeader title="Submissions" subtitle={`${sorted.length} record${sorted.length === 1 ? '' : 's'}`} />
      <div style={{ marginBottom: 18 }}>
        <select value={filterTemplate} onChange={e => setFilterTemplate(e.target.value)} style={selectStyle(C)}>
          <option value="All">All templates</option>
          {templates.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
        </select>
      </div>
      <div style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, overflow: 'hidden' }}>
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1.2fr 1fr 1fr 1fr', padding: '10px 16px', fontSize: 11, fontWeight: 600, color: C.faint, textTransform: 'uppercase', letterSpacing: '0.04em', borderBottom: `1px solid ${C.border}`, background: C.bg }}>
          <span>Entry</span><span>Template</span><span>Filled By</span><span>Date</span><span>Department</span>
        </div>
        {sorted.map(s => (
          <button key={s.id} onClick={() => onView(s)} style={{ display: 'grid', gridTemplateColumns: '2fr 1.2fr 1fr 1fr 1fr', padding: '13px 16px', fontSize: 13.5, borderBottom: `1px solid ${C.border}`, alignItems: 'center', width: '100%', border: 'none', background: 'none', cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit' }}>
            <span style={{ fontWeight: 600, color: C.navy }}>{s.title || s.templateName}</span>
            <span style={{ color: C.dim }}>{s.templateName}</span>
            <span style={{ color: C.dim }}>{s.filledBy}</span>
            <span style={{ color: C.dim, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>{s.filledAt}</span>
            <span style={{ color: C.dim }}>{s.department || '—'}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ================= Small shared components =================
function PageHeader({ title, subtitle }) {
  return (
    <div style={{ marginBottom: 22 }}>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, margin: 0, color: C.navy }}>{title}</h1>
      {subtitle && <div style={{ fontSize: 13.5, color: C.dim, marginTop: 4 }}>{subtitle}</div>}
    </div>
  );
}
function KpiCard({ label, value, color }) {
  return (
    <div style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, padding: '16px 18px', borderTop: `3px solid ${color}` }}>
      <div style={{ fontSize: 26, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: C.text }}>{value}</div>
      <div style={{ fontSize: 12, color: C.dim, marginTop: 4 }}>{label}</div>
    </div>
  );
}
function ChartCard({ title, children }) {
  return (
    <div style={{ background: '#fff', border: `1px solid ${C.border}`, borderRadius: 8, padding: '18px 20px' }}>
      <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>{title}</div>
      {children}
    </div>
  );
}
function IconBtn({ children, onClick, title }) {
  return (
    <button onClick={onClick} title={title} type="button" style={{ border: `1px solid ${C.border}`, background: '#fff', borderRadius: 5, width: 28, height: 28, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', color: C.dim }}>
      {children}
    </button>
  );
}
function EmptyState({ title, body, onNew, canCreate = true }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', background: '#fff', border: `1px dashed ${C.borderStrong}`, borderRadius: 10, padding: '48px 40px', maxWidth: 480 }}>
      <FileText size={28} color={C.faint} />
      <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 19, margin: '16px 0 6px' }}>{title}</h2>
      <p style={{ fontSize: 13.5, color: C.dim, margin: '0 0 18px', lineHeight: 1.6 }}>{body}</p>
      {canCreate && <button onClick={onNew} style={btnPrimary(C)}><PlusCircle size={14} /> New Document</button>}
    </div>
  );
}
function Modal({ children, onClose, narrow }) {
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(11,32,51,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, padding: 20 }} onClick={onClose}>
      <div className="modal-inner" style={{ background: '#fff', borderRadius: 10, padding: 28, width: '100%', maxWidth: narrow ? 400 : 660, maxHeight: '84vh', overflow: 'auto', position: 'relative', boxShadow: '0 20px 60px rgba(0,0,0,0.25)' }} onClick={e => e.stopPropagation()}>
        <button onClick={onClose} style={{ position: 'absolute', top: 16, right: 16, border: 'none', background: 'none', cursor: 'pointer', color: C.faint }}><X size={18} /></button>
        {children}
      </div>
    </div>
  );
}
function selectStyle(C) { return { padding: '9px 12px', borderRadius: 5, border: `1px solid ${C.border}`, fontSize: 13.5, background: '#fff', fontFamily: 'inherit' }; }
function btnPrimary(C) { return { display: 'inline-flex', alignItems: 'center', gap: 7, background: C.green, color: '#fff', border: 'none', borderRadius: 6, padding: '10px 18px', fontSize: 13.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit' }; }
function btnGhost(C) { return { display: 'inline-flex', alignItems: 'center', gap: 7, background: '#fff', color: C.text, border: `1px solid ${C.border}`, borderRadius: 6, padding: '10px 18px', fontSize: 13.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit' }; }
