import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Box,
  Typography,
  Card,
  CardContent,
  List,
  ListItem,
  ListItemText,
  IconButton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Button,
  Snackbar,
  Select,
  MenuItem,
  Menu,
  Slide,
  Chip,
  Tooltip,
  Divider,
  InputAdornment,
  Autocomplete,
} from "@mui/material";
import { DataGrid, GridToolbarQuickFilter } from "@mui/x-data-grid";
import EditIcon from "@mui/icons-material/Edit";
import DeleteIcon from "@mui/icons-material/Delete";
import CloseIcon from "@mui/icons-material/Close";
import SendIcon from "@mui/icons-material/Send";
import AttachFileIcon from "@mui/icons-material/AttachFile";
import ContentPasteSearchIcon from '@mui/icons-material/ContentPasteSearch';
import MenuIcon from '@mui/icons-material/Menu';

/* ========= COPY/PASTE FIX =========
   In production, ignore VITE_API_URL and use same-origin ('').
   In dev, use VITE_API_URL or fall back to http://localhost:5000.
*/
const API_BASE = import.meta.env.PROD
  ? "" 
  : (import.meta.env.VITE_API_URL || "http://localhost:5000").trim().replace(/\/+$/, "");
if (import.meta.env.DEV) console.log("API_BASE ->", API_BASE);
/* ================================== */

const STATUS_OPTIONS = ["New", "Matched", "Posting", "Completed", "In Query"];
const FOLDERS = ["UK", "Ireland", "Foreign", "Spain", "GmbH"];
const ASSIGNEES = ["Allan Perry", "Marcos Silva", "Caroline Stathatos"];

// Optional: one or more shared mailboxes to send from
const SHARED_MAILBOXES = (
  import.meta.env.VITE_SHARED_MAILBOXES?.split(",") || [
    "accounts.queries@gear4music.com",
  ]
).map((s) => s.trim()).filter(Boolean);

const ORDER = ["New", "Matched", "Posting", "Completed"];

const formatUKDate = (isoString) => {
  if (!isoString) return "";
  const date = new Date(isoString);
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
};

// Extract unique G-numbers (e.g. G123456) from an array of note objects
const extractGNumbers = (notes) => {
  if (!Array.isArray(notes) || notes.length === 0) return [];
  const rx = /\bG(\d{6})\b/gi;
  const seen = new Set();
  const out = [];
  for (const n of notes) {
    if (!n || !n.text) continue;
    let m;
    while ((m = rx.exec(n.text))) {
      const g = `G${m[1]}`.toUpperCase();
      if (!seen.has(g)) {
        seen.add(g);
        out.push(g);
      }
    }
  }
  return out;
};

// Normalize supplier string for grouping/filtering (trim + case-fold)
const normalizeSupplier = (s) => String(s || '').trim().toLowerCase();

function SlideUpTransition(props) {
  return <Slide {...props} direction="up" />;
}

function ExportMenu({ onExport, selectedCount }) {
  const [anchor, setAnchor] = useState(null);
  return (
    <>
      <IconButton size="small" onClick={(e) => setAnchor(e.currentTarget)}>
        <MenuIcon />
      </IconButton>
      <Menu anchorEl={anchor} open={Boolean(anchor)} onClose={() => setAnchor(null)}>
        <MenuItem onClick={() => { setAnchor(null); onExport && onExport(); }}>Export CSV</MenuItem>
        <MenuItem onClick={() => { setAnchor(null); window.open('https://gear4music.sharepoint.com/sites/Gear4musicIntranetPortal/AIProcessing/Forms/AllItems.aspx', '_blank', 'noopener,noreferrer'); }}>
          Open Sharepoint
        </MenuItem>
      </Menu>
    </>
  )
}

/** Thin connector line between pills — now colourable & fill-aware */
const Connector = ({ filled = false, color = "divider" }) => (
  <Box
    sx={{
      width: 40,
      height: 2,
      bgcolor: filled ? color : "divider",
      mx: 1,
      borderRadius: 1,
      transition: "background-color 200ms ease",
    }}
  />
);

/** Small pill button for progress bar — disabled style now respects `filled` */
const Pill = ({ label, color, filled, disabled, onClick }) => (
  <Button
    size="small"
    onClick={onClick}
    disabled={disabled}
    disableElevation
    sx={{
      borderRadius: 999,
      px: 2,
      py: 0.5,
      fontSize: 12,
      lineHeight: 1,
      bgcolor: filled ? color : "transparent",
      color: filled ? "#fff" : color,
      border: "1px solid",
      borderColor: color,
      transition: "background-color 200ms ease, color 200ms ease, border-color 200ms ease",
      "&:hover": {
        bgcolor: filled ? color : "rgba(0,0,0,0.04)",
        borderColor: color,
      },
      "&.Mui-disabled": {
        bgcolor: filled ? color : "action.disabledBackground",
        color: filled ? "#fff" : "text.disabled",
        borderColor: filled ? color : "divider",
        opacity: 1,
      },
    }}
  >
    {label}
  </Button>
);

/**
 * Custom toolbar in the DataGrid:
 * Left = quick search
 * Center = pill progress bar
 * Right = actions (Open / Edit / Delete)
 */
function ProcessToolbar(props) {
  const {
    statusNow,
    canMatch,
    canPost,
    canComplete,
    startMatch,
    startPost,
    startComplete,
    isFilled,

    // NEW: selected + handlers
    selectedCount,
    onOpenSelected,
    onEditSelected,
    onDeleteSelected,
    // export handler
    onExport,
  } = props;

  // helper: connector fills once we've reached the *right-hand* step
  const connTo = (toLabel) => ORDER.indexOf(toLabel) <= ORDER.indexOf(statusNow);

  return (
    <Box sx={{ display: "flex", alignItems: "center", gap: 2, px: 1, py: 0.5 }}>
      {/* Left: Quick search */}
      <GridToolbarQuickFilter />

      {/* Center: pill buttons + thin connectors */}
      <Box sx={{ flex: 1, display: "flex", justifyContent: "center" }}>
        <Box
          sx={{
            display: "flex",
            alignItems: "center",
            px: 1,
            py: 0.5,
            borderRadius: 999,
            border: "1px solid",
            borderColor: "divider",
            backgroundColor: "transparent",
          }}
        >
          <Pill label="New" color="#e53935" filled={isFilled("New")} disabled />
          <Connector filled={connTo("Matched")} color="#e53935" />
          <Pill
            label="Match"
            color="#8e24aa"
            filled={isFilled("Matched")}
            disabled={!canMatch}
            onClick={startMatch}
          />
          <Connector filled={connTo("Posting")} color="#8e24aa" />
          <Pill
            label="Post"
            color="#1e88e5"
            filled={isFilled("Posting")}
            disabled={!canPost}
            onClick={startPost}
          />
          <Connector filled={connTo("Completed")} color="#1e88e5" />
          <Pill
            label="Complete"
            color="#43a047"
            filled={isFilled("Completed")}
            disabled={!canComplete}
            onClick={startComplete}
          />
        </Box>
      </Box>

      {/* Right: Action buttons (multi-select aware) */}
      <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
        <Button
          size="small"
          variant="outlined"
          onClick={onOpenSelected}
          disabled={selectedCount === 0}
        >
          Open
        </Button>
        <Button
          size="small"
          variant="outlined"
          onClick={onEditSelected}
          disabled={selectedCount === 0}
          startIcon={<EditIcon fontSize="small" />}
        >
          Edit
        </Button>
        <Button
          size="small"
          color="error"
          variant="outlined"
          onClick={onDeleteSelected}
          disabled={selectedCount === 0}
          startIcon={<DeleteIcon fontSize="small" />}
        >
          Delete
        </Button>
        {/* Burger menu (to the right of Delete) */}
        <Box>
          <ExportMenu onExport={onExport} selectedCount={selectedCount} />
        </Box>
      </Box>
    </Box>
  );
}

function InvoiceDashboard() {
  const [rows, setRows] = useState([]);
  const [filterSupplier, setFilterSupplier] = useState(null);
  const [supplierQuery, setSupplierQuery] = useState('');

  const [selectedRow, setSelectedRow] = useState(null);
  const [rowSelectionModel, setRowSelectionModel] = useState([]);

  const [selectedNote, setSelectedNote] = useState(null);
  const [noteDialogOpen, setNoteDialogOpen] = useState(false);
  const [noteText, setNoteText] = useState("");
  const [noteReadOnly, setNoteReadOnly] = useState(false);
  const [deleteDialog, setDeleteDialog] = useState({ open: false, ids: [], message: '' });

  // Status column header filter
  const [statusFilter, setStatusFilter] = useState("All");

  // New items pop-up
  const [newItemsCount, setNewItemsCount] = useState(0);
  const [snackbarOpen, setSnackbarOpen] = useState(false);
  const [editDialogOpen, setEditDialogOpen] = useState(false);
  const [editForm, setEditForm] = useState({ hub: '', folder: FOLDERS[0], assigned: ASSIGNEES[0], ref: '' , status: 'New'});

  // Burger menu
  const [menuAnchor, setMenuAnchor] = useState(null);
  const openMenu = (e) => setMenuAnchor(e.currentTarget);
  const closeMenu = () => setMenuAnchor(null);

  const downloadCsv = (rowsToExport) => {
    if (!rowsToExport || !rowsToExport.length) return;
    const fields = ['id','supplier','hub','type','invoiceno','invoice_date','po','folder','assigned','ref','last_modified','created_on','status'];
    const csv = [fields.join(',')].concat(rowsToExport.map(r => {
      return fields.map(f => {
        const v = r[f];
        if (v === null || v === undefined) return '';
        return String(v).replace(/"/g,'""');
      }).map(x => `"${x}"`).join(',')
    })).join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `invoices-${new Date().toISOString().slice(0,10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  const prevIdsRef = useRef(new Set());

  // Process dialogs
  const [folderDialogOpen, setFolderDialogOpen] = useState(false);
  const [chosenFolder, setChosenFolder] = useState(FOLDERS[0]);

  const [assigneeDialogOpen, setAssigneeDialogOpen] = useState(false);
  const [chosenAssignee, setChosenAssignee] = useState(ASSIGNEES[0]);

  const [refDialogOpen, setRefDialogOpen] = useState(false);
  const [refValue, setRefValue] = useState("");

  // ---------- EMAIL: Outlook-like compose ----------
  const [emailDialogOpen, setEmailDialogOpen] = useState(false);
  const [emailFrom, setEmailFrom] = useState(SHARED_MAILBOXES[0] || "");
  const [emailTo, setEmailTo] = useState("");
  const [emailCc, setEmailCc] = useState("");
  const [emailBcc, setEmailBcc] = useState("");
  const [emailSubject, setEmailSubject] = useState("");
  const [emailBody, setEmailBody] = useState("");
  const [emailFiles, setEmailFiles] = useState([]); // File[]
  const [isSending, setIsSending] = useState(false);
  // For people search autocomplete
  const [peopleOptions, setPeopleOptions] = useState([]);
  // keep lastTerm, abort controller and debounce timer in a ref
  const peopleFetchRef = useRef({ lastTerm: '', abort: null, timer: null });
  const [peopleSearchError, setPeopleSearchError] = useState(null);

  const supplierCounts = useMemo(() => {
    // key: normalized supplier -> { name: displayName, count }
    const map = {};
    for (const row of rows) {
      const raw = row.supplier || '';
      const key = normalizeSupplier(raw);
      if (!map[key]) map[key] = { name: raw.trim() || raw, count: 0 };
      map[key].count += 1;
    }
    // Convert to plain object mapping displayName -> count but keep normalized keys available
    const out = {};
    for (const [k, v] of Object.entries(map)) {
      out[k] = { name: v.name, count: v.count };
    }
    return out;
  }, [rows]);

  // Detect duplicate rows: same supplier + same invoice number (case-insensitive supplier)
  const duplicateRowIds = useMemo(() => {
    const map = new Map();
    for (const r of rows) {
      const key = `${normalizeSupplier(r.supplier)}||${String(r.invoiceno || '').trim()}`;
      const arr = map.get(key) || [];
      arr.push(r.id);
      map.set(key, arr);
    }
    const dupIds = new Set();
    for (const arr of map.values()) {
      if (arr.length > 1) arr.forEach(id => dupIds.add(id));
    }
    return dupIds;
  }, [rows]);

  // (removed supplier-wide G-number summary — G numbers are shown only per-invoice)

  // Polling invoices
  useEffect(() => {
    const fetchInvoices = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/invoices`);
        const data = await res.json();

        const prevIds = prevIdsRef.current;
        const newRows = data.map((row) => ({
          ...row,
          isNew: !prevIds.has(row.id),
        }));

        const newCount = newRows.filter((r) => r.isNew).length;
        if (newCount > 0) {
          setNewItemsCount(newCount);
          setSnackbarOpen(true);
        }

        prevIdsRef.current = new Set(data.map((r) => r.id));
        setRows(newRows);

        newRows
          .filter((r) => r.isNew)
          .forEach((r) => {
            setTimeout(() => {
              setRows((prev) =>
                prev.map((row) =>
                  row.id === r.id ? { ...row, isNew: false } : row
                )
              );
            }, 4000);
          });
      } catch (err) {
        console.error("Error fetching invoices:", err);
      }
    };

    fetchInvoices();
    const interval = setInterval(fetchInvoices, 5000);
    return () => clearInterval(interval);
  }, []);

  // Filter + selection sync
  const filteredRows = useMemo(() => {
    let base = filterSupplier ? rows.filter((r) => normalizeSupplier(r.supplier) === filterSupplier) : rows;
    if (statusFilter !== "All") {
      if (statusFilter === 'In Query') {
        base = base.filter((r) => Array.isArray(r.notes) && r.notes.length > 0);
      } else {
        base = base.filter((r) => (r.status || "New") === statusFilter);
      }
    }
    return base;
  }, [rows, filterSupplier, statusFilter]);

  useEffect(() => {
    if (rowSelectionModel.length === 0) {
      setSelectedRow(null);
      return;
    }
    const id = rowSelectionModel[0];
    const source = filterSupplier ? filteredRows : rows;
    const match = source.find((r) => r.id === id) || null;
    setSelectedRow(match);
  }, [rows, filteredRows, filterSupplier, rowSelectionModel]);

  // Optimistic patch (+ server persistence via PATCH /api/invoices/:id)
  const applyPatch = async (invoiceId, patch) => {
    setRows((prev) => prev.map((r) => (r.id === invoiceId ? { ...r, ...patch } : r)));
    setSelectedRow((prev) => (prev && prev.id === invoiceId ? { ...prev, ...patch } : prev));
    try {
      await fetch(`${API_BASE}/api/invoices/${invoiceId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(patch),
      });
    } catch {
      // If the backend route isn't present yet, it will simply remain optimistic in UI.
    }
  };

  // Current status helpers (fill pills up to current stage)
  const statusNow = selectedRow?.status || "New";
  const currentIdx = ORDER.indexOf(statusNow);
  const isFilled = (label) => ORDER.indexOf(label) <= currentIdx;

  // Button enabling
  const canMatch = !!selectedRow && statusNow === "New";
  const canPost = !!selectedRow && statusNow === "Matched";
  const canComplete = !!selectedRow && statusNow === "Posting";

  // Process actions
  const startMatch = () => {
    if (!selectedRow) return;
    setChosenFolder(
      selectedRow.folder && FOLDERS.includes(selectedRow.folder)
        ? selectedRow.folder
        : FOLDERS[0]
    );
    setFolderDialogOpen(true);
  };
  const confirmMatch = async () => {
    if (!selectedRow) return;
    await applyPatch(selectedRow.id, {
      status: "Matched",
      folder: chosenFolder,
      last_modified: new Date().toISOString(),
    });
    setFolderDialogOpen(false);
  };

  const startPost = () => {
    if (!selectedRow) return;
    setChosenAssignee(
      selectedRow.assigned && ASSIGNEES.includes(selectedRow.assigned)
        ? selectedRow.assigned
        : ASSIGNEES[0]
    );
    setAssigneeDialogOpen(true);
  };
  const confirmPost = async () => {
    if (!selectedRow) return;
    await applyPatch(selectedRow.id, {
      status: "Posting",
      assigned: chosenAssignee,
      last_modified: new Date().toISOString(),
    });
    setAssigneeDialogOpen(false);
  };

  const startComplete = () => {
    if (!selectedRow) return;
    setRefValue("");
    setRefDialogOpen(true);
  };
  const confirmComplete = async () => {
    if (!selectedRow) return;
    if (!/^\d{6}$/.test(refValue)) return;
    await applyPatch(selectedRow.id, {
      status: "Completed",
      ref: refValue,
      last_modified: new Date().toISOString(),
    });
    setRefDialogOpen(false);
  };

  // Notes
  const handleViewNote = (note) => {
    if (!selectedRow) return;
    setSelectedNote(note);
    setNoteText(note.text);
    setNoteReadOnly(true);
    setNoteDialogOpen(true);
  };

  const handleEditNote = (note) => {
    if (!selectedRow) return;
    setSelectedNote(note);
    setNoteText(note.text);
    setNoteReadOnly(false);
    setNoteDialogOpen(true);
  };

  const handleSaveNote = async () => {
    if (!selectedRow) return;
    const invoiceId = selectedRow.id;

    if (selectedNote) {
      try {
        const res = await fetch(
          `${API_BASE}/api/invoices/${invoiceId}/notes/${selectedNote.id}`,
          {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: noteText, supplier: selectedRow.supplier }),
          }
        );
        const updatedNote = await res.json();
        const updatedRow = {
          ...selectedRow,
          notes: (Array.isArray(selectedRow.notes) ? selectedRow.notes : []).map((n) =>
            n.id === updatedNote.id ? updatedNote : n
          ),
        };
        setRows((prev) => prev.map((r) => (r.id === invoiceId ? updatedRow : r)));
        setSelectedRow(updatedRow);
      } catch (err) {
        console.error(err);
      }
    } else {
      try {
        const newNotePayload = {
          text: noteText,
          date: new Date().toISOString().split("T")[0],
          supplier: selectedRow.supplier,
        };
        const res = await fetch(`${API_BASE}/api/invoices/${invoiceId}/notes`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(newNotePayload),
        });
        const savedNote = await res.json();
        const updatedRow = {
          ...selectedRow,
          notes: Array.isArray(selectedRow.notes) ? [...selectedRow.notes, savedNote] : [savedNote],
        };
        setRows((prev) => prev.map((r) => (r.id === invoiceId ? updatedRow : r)));
        setSelectedRow(updatedRow);
      } catch (err) {
        console.error(err);
      }
    }

    setNoteDialogOpen(false);
    setSelectedNote(null);
    setNoteText("");
  };

  const handleDeleteNote = async (noteId) => {
    if (!selectedRow) return;
    const invoiceId = selectedRow.id;
    if (!window.confirm("Are you sure you want to delete this note?")) return;
    try {
      await fetch(`${API_BASE}/api/invoices/${invoiceId}/notes/${noteId}`, { method: "DELETE" });
      const updatedRow = {
        ...selectedRow,
        notes: (Array.isArray(selectedRow.notes) ? selectedRow.notes : []).filter((n) => n.id !== noteId),
      };
      setRows((prev) => prev.map((r) => (r.id === invoiceId ? updatedRow : r)));
      setSelectedRow(updatedRow);
    } catch (err) {
      console.error(err);
    }
  };

  /** ---------- NEW: Multi-select toolbar actions ---------- */
  const getSelectedRows = useCallbackIdsToRows(rowSelectionModel, rows);

  const handleOpenSelected = async () => {
    const selected = getSelectedRows();
    if (!selected.length) return;

    // open placeholders synchronously (no noopener/noreferrer)
    const tabs = selected.map(() => {
      const w = window.open('about:blank', '_blank');
      if (w) {
        w.document.write('<p style="font-family:system-ui;margin:16px">Opening PDF…</p>');
        w.document.close();
      }
      return w;
    });

    await Promise.all(
      selected.map(async (r, i) => {
        const w = tabs[i];
        try {
          const resp = await fetch(`${API_BASE}/api/invoices/${r.id}/blob-url`);
          if (!resp.ok) throw new Error(await resp.text());
          const { url } = await resp.json();
          if (!url) throw new Error('No url in response');

          // navigate the already-opened tab
          if (w) w.location.href = url; // use href (most compatible)
        } catch (err) {
          console.error('Open failed for', r.id, err);
          if (w && !w.closed) {
            w.document.body.innerHTML =
              `<div style="font-family:system-ui;margin:16px">
                 <div style="margin-bottom:8px">Couldn’t open this file.</div>
                 <code style="font-size:12px;white-space:pre-wrap">${String(err.message || err)}</code>
               </div>`;
          }
        }
      })
    );
  };

  const handleEditSelected = () => {
    const selected = getSelectedRows();
    if (selected.length === 0) return;
    // If single row selected, prefill with that row's values
    if (selected.length === 1) {
      const r = selected[0];
      setEditForm({
        hub: r.hub || '',
        folder: r.folder || FOLDERS[0],
        assigned: r.assigned || ASSIGNEES[0],
        ref: r.ref || '',
        status: r.status || 'New'
      });
      setEditDialogOpen(true);
      return;
    }

    // Multiple rows: clear form (user chooses values to apply)
    setEditForm({ hub: '', folder: FOLDERS[0], assigned: ASSIGNEES[0], ref: '', status: 'New' });
    setEditDialogOpen(true);
  };

  const handleDeleteSelected = () => {
    const selected = getSelectedRows();
    if (selected.length === 0) return;

    const message =
      selected.length === 1
        ? `Delete invoice ${selected[0].id}?`
        : `Delete ${selected.length} selected invoices?`;

    const ids = selected.map(r => r.id);
    setDeleteDialog({ open: true, ids, message });
  };

  const cancelDeleteSelected = () => {
    setDeleteDialog({ open: false, ids: [], message: '' });
  };

  const confirmDeleteSelected = async () => {
    const ids = deleteDialog.ids;
    if (!ids.length) {
      cancelDeleteSelected();
      return;
    }

    setRows(prev => prev.filter(r => !ids.includes(r.id)));
    setRowSelectionModel([]);

    cancelDeleteSelected();

    try {
      await Promise.allSettled(
        ids.map(id =>
          fetch(`${API_BASE}/api/invoices/${id}`, { method: 'DELETE' })
        )
      );
    } catch (e) {
      console.error('Delete error:', e);
    }
  };
  /** ------------------------------------------------------- */

  // -------- EMAIL helpers --------
  const openEmailComposer = async () => {
    if (!selectedRow) return;

    // default subject: "<Supplier> <InvoiceNo>"
    const subject = `${selectedRow.supplier || ""} ${selectedRow.invoiceno || ""}`.trim();
    setEmailSubject(subject);
    setEmailBody("");
    setEmailTo("");
    setEmailCc("");
    setEmailBcc("");
    setEmailFiles([]);
    setEmailFrom(SHARED_MAILBOXES[0] || "");

    setEmailDialogOpen(true);
  };

  const addFiles = (fileList) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    setEmailFiles(prev => {
      // de-dup by name+size
      const key = (f) => `${f.name}:${f.size}`;
      const existing = new Set(prev.map(key));
      const merged = [...prev];
      for (const f of files) if (!existing.has(key(f))) merged.push(f);
      return merged;
    });
  };

  const removeFile = (idx) => {
    setEmailFiles((prev) => prev.filter((_, i) => i !== idx));
  };

  const attachInvoicePdf = async () => {
    if (!selectedRow) return;
    try {
      // Use the blob-url helper to fetch the PDF and turn it into a File
      const resp = await fetch(`${API_BASE}/api/invoices/${selectedRow.id}/blob-url`);
      const { url } = await resp.json();
      if (!url) throw new Error("No blob url returned");
      const blob = await fetch(url).then((r) => r.blob());
      const suggestedName = `${selectedRow.supplier || "invoice"}-${selectedRow.invoiceno || selectedRow.id}.pdf`;
      const file = new File([blob], suggestedName, { type: blob.type || "application/pdf" });
      addFiles([file]);
    } catch (e) {
      alert(`Could not attach invoice PDF: ${e?.message || e}`);
    }
  };

  const handleSendEmail = async () => {
    if (!selectedRow) return;
    if (!emailTo.trim()) {
      alert("Please enter at least one recipient (To).");
      return;
    }
    setIsSending(true);

    try {
      const form = new FormData();
      form.append("from", emailFrom);
      form.append("to", emailTo);
      if (emailCc) form.append("cc", emailCc);
      if (emailBcc) form.append("bcc", emailBcc);
      form.append("subject", emailSubject || "");
      form.append("body", emailBody || "");
      form.append("invoiceId", selectedRow.id);

      emailFiles.forEach((f, i) => form.append("attachments", f, f.name));

      const resp = await fetch(`${API_BASE}/api/email/send`, {
        method: "POST",
        body: form,
      });

      if (!resp.ok) throw new Error(await resp.text());
      const result = await resp.json();

      // Add the email body to Notes (as requested)
      try {
        const newNotePayload = {
          text: `Email sent${emailTo ? ` to ${emailTo}` : ""}${emailCc ? `, cc ${emailCc}` : ""}${emailBcc ? `, bcc ${emailBcc}` : ""}\n\nSubject: ${emailSubject}\n\n${emailBody}`,
          date: new Date().toISOString().split("T")[0],
          supplier: selectedRow.supplier,
        };
        const res = await fetch(`${API_BASE}/api/invoices/${selectedRow.id}/notes`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(newNotePayload),
        });
        const savedNote = await res.json();
        const updatedRow = {
          ...selectedRow,
          notes: Array.isArray(selectedRow.notes) ? [...selectedRow.notes, savedNote] : [savedNote],
        };
        setRows((prev) => prev.map((r) => (r.id === selectedRow.id ? updatedRow : r)));
        setSelectedRow(updatedRow);
      } catch (err) {
        console.error("Failed to save email body to notes:", err);
      }

      setEmailDialogOpen(false);
      setEmailFiles([]);
    } catch (e) {
      alert(`Send failed: ${e?.message || e}`);
    } finally {
      setIsSending(false);
    }
  };

  // Columns (status header dropdown remains)
  const columns = useMemo(
    () => [
      {
        field: "supplier",
        headerName: "Supplier",
        flex: 1.6,
        renderCell: (params) => {
          const supplier = params.value || '';
          // Only show G-numbers that belong to this invoice (row-local)
          const rowG = extractGNumbers(params.row.notes || []);
          const gText = rowG.length ? ` [${rowG.join(', ')}]` : '';
          return <Typography sx={{ fontSize: 13 }}>{`${supplier}${gText}`}</Typography>;
        }
      },
      { field: "hub", headerName: "Hub", flex: 0.75 },
      { field: "type", headerName: "Type", flex: 0.75 },
      { field: "invoiceno", headerName: "Invoice No", flex: 1 },

      {
        field: "invoice_date",
        headerName: "Invoice Date",
        flex: 0.6,
        valueGetter: (params) => formatUKDate(params.row.invoice_date),
      },
      { field: "po", headerName: "PO", flex: 1 },
      { field: "folder", headerName: "Folder", flex: 1 },
      { field: "assigned", headerName: "Assigned", flex: 1 },
      { field: "ref", headerName: "Ref", flex: 1 },
      {
        field: "last_modified",
        headerName: "Last Modified",
        flex: 0.6,
        valueGetter: (params) => formatUKDate(params.row.last_modified),
      },
      {
        field: "created_on",
        headerName: "Created On",
        flex: 0.6,
        valueGetter: (params) => formatUKDate(params.row.created_on),
      },
      {
        field: "status",
        headerName: "Status",
        flex: 1,
        valueGetter: (params) => params.row.status || "New",
        renderCell: (params) => {
          return <Typography sx={{ fontSize: 13 }}>{params.value}</Typography>
        },
        renderHeader: () => (
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            <Typography variant="caption" sx={{ fontWeight: 600 }}>
              Status
            </Typography>
            <Select
              size="small"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              displayEmpty
              sx={{
                fontSize: 12,
                height: 26,
                "& .MuiSelect-select": { py: 0.2, minHeight: 0 },
              }}
            >
              <MenuItem value="All">All</MenuItem>
              {STATUS_OPTIONS.map((s) => (
                <MenuItem key={s} value={s}>
                  {s}
                </MenuItem>
              ))}
            </Select>
          </Box>
        ),
      },
      {
        field: 'in_query',
        headerName: '',
        width: 48,
        sortable: false,
        filterable: false,
        align: 'center',
        headerAlign: 'center',
        renderCell: (params) => {
          const hasNotes = Array.isArray(params.row.notes) && params.row.notes.length > 0;
          return hasNotes ? (
            <ContentPasteSearchIcon sx={{ color: 'error.main', fontSize: 18 }} />
          ) : null
        }
      },
    ],
    [statusFilter]
  );

  return (
    <Box sx={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <Typography variant="h5" gutterBottom>
        Invoice Manager
      </Typography>
      {/* top-right menu removed — menu is now inside the DataGrid toolbar */}

      {/* DataGrid with custom toolbar */}
      <Box sx={{ flexGrow: 1, minHeight: 0 }}>
        <DataGrid
          rows={filteredRows}
          columns={columns}
          getRowClassName={(params) => (duplicateRowIds.has(params.id) ? 'dup-row' : '')}
          rowHeight={30}
          headerHeight={28}
          density="compact"
          checkboxSelection
          disableRowSelectionOnClick
          rowSelectionModel={rowSelectionModel}
          onRowSelectionModelChange={(newModel) => setRowSelectionModel(newModel)}
          onCellClick={(params) => {
            if (params.field === "supplier") {
              setRowSelectionModel([params.id]);
              setSelectedRow(params.row);
            }
          }}
          slots={{ toolbar: ProcessToolbar }}
          slotProps={{
            toolbar: {
              statusNow,
              canMatch,
              canPost,
              canComplete,
              startMatch,
              startPost,
              startComplete,
              isFilled,

              // NEW: actions + selection count
              selectedCount: rowSelectionModel.length,
              onOpenSelected: handleOpenSelected,
              onEditSelected: handleEditSelected,
              onDeleteSelected: handleDeleteSelected,
              onExport: async () => {
                const ids = rowSelectionModel.length ? rowSelectionModel.join(',') : '';
                const url = ids ? `${API_BASE}/api/invoices/export?ids=${ids}` : `${API_BASE}/api/invoices/export`;
                try {
                  const resp = await fetch(url);
                  if (!resp.ok) throw new Error(await resp.text());
                  const blob = await resp.blob();
                  const dlUrl = URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.href = dlUrl;
                  a.download = `invoices-${new Date().toISOString().slice(0,10)}.csv`;
                  a.click();
                  URL.revokeObjectURL(dlUrl);
                } catch (e) {
                  alert('Export failed: ' + (e?.message || e));
                }
              }
            },
          }}
          sx={{
            fontSize: 13,
            "& .MuiDataGrid-columnHeaders": { fontSize: 12, color: "#68adf1ff" },
            "& .MuiDataGrid-cell": {
              py: 0,
              lineHeight: "14px",
              display: "flex",
              alignItems: "center",
            },
            "& .dup-row": {
              backgroundColor: 'rgba(255,0,0,0.28)',
              borderLeft: '4px solid rgba(255,0,0,0.6)'
            },
            "& .MuiDataGrid-row.Mui-selected, .MuiDataGrid-row.Mui-selected:hover": {
              backgroundColor: "rgba(25,118,210,0.2)",
            },
            "& .MuiCheckbox-root": { width: "10px", height: "10px", padding: 0 },
            "& .MuiCheckbox-root svg": { transform: "scale(0.5)", transformOrigin: "center" },
          }}
        />
      </Box>

      {/* Bottom panels (Suppliers + Notes) */}
      <Box sx={{ display: "flex", mt: 1, flexShrink: 0 }}>
        <Card sx={{ flex: 1, mr: 1, height: "310px", position: "relative" }}>
          <Box
            display="flex"
            justifyContent="space-between"
            alignItems="center"
            sx={{ position: "sticky", top: 0, background: "transparent", zIndex: 1, p: 1, pb: 0.5 }}
          >
            <Typography variant="body2" sx={{ fontSize: 18 }}>Suppliers</Typography>
            <Button size="small" sx={{ fontSize: 14, textTransform: "none" }} onClick={() => { setFilterSupplier(null); setSupplierQuery(''); }}>
              Show All
            </Button>
          </Box>
          <Box sx={{ overflowY: "auto", height: "calc(100% - 32px)", px: 1 }}>
            <Box sx={{ px: 0.5, pb: 0.5 }}>
              <TextField
                size="small"
                placeholder="Search suppliers..."
                value={supplierQuery}
                onChange={(e) => setSupplierQuery(e.target.value)}
                fullWidth
              />
            </Box>
            <List dense>
              {Object.entries(supplierCounts)
                .filter(([norm, obj]) => !supplierQuery || obj.name.toLowerCase().includes(supplierQuery.toLowerCase()))
                .map(([norm, obj]) => (
                  <ListItem key={norm} button onClick={() => setFilterSupplier(norm)} sx={{ py: 0.3 }}>
                    <ListItemText
                      primaryTypographyProps={{ fontSize: 14, color: "inherit" }}
                      primary={`${obj.name} (${obj.count})`}
                    />
                  </ListItem>
                ))}
            </List>
          </Box>
        </Card>

        <Card sx={{ flex: 3, height: "310px"}}>
          <CardContent sx={{ p: 1, height: "100%", overflowY: "auto" }}>
            <Box display="flex" justifyContent="space-between" alignItems="center" mb={0.5}>
              <Typography variant="body2" sx={{ fontSize: 16 }}>Notes</Typography>
              <Box sx={{ display: 'flex', gap: 1 }}>
                <Button
                  size="small"
                  sx={{ fontSize: 14, textTransform: "none" }}
                  onClick={() => { setSelectedNote(null); setNoteText(""); setNoteReadOnly(false); setNoteDialogOpen(true); }}
                  disabled={!selectedRow}
                >
                  + Add Note
                </Button>
                {/* NEW: Email button next to + Add Note */}
                <Button
                  size="small"
                  startIcon={<SendIcon fontSize="small" />}
                  sx={{
                    fontSize: 14,
                    textTransform: "none",
                    // keep the icon compact so the button visually matches the + Add Note button
                    '& .MuiButton-startIcon': { mr: 0.5 },
                    px: 1.5,
                  }}
                  onClick={openEmailComposer}
                  disabled={!selectedRow}
                >
                  Email
                </Button>
              </Box>
            </Box>
                <List dense>
              {selectedRow && Array.isArray(selectedRow.notes) && selectedRow.notes.length > 0 ? (
                selectedRow.notes.map((note) => (
                  <ListItem
                    key={note.id}
                    button
                    onClick={() => handleViewNote(note)}
                    sx={{ py: 0.3, display: "flex", alignItems: "center" }}
                  >
                    <ListItemText
                      primaryTypographyProps={{ fontSize: 14 }}
                      sx={{ flex: 1, mr: 1, whiteSpace: 'normal' }}
                      primary={`${note.text} (${note.date})`}
                    />
                    <Box>
                      <IconButton size="small" onClick={(e) => { e.stopPropagation(); handleEditNote(note); }}><EditIcon fontSize="inherit" /></IconButton>
                      <IconButton size="small" onClick={(e) => { e.stopPropagation(); handleDeleteNote(note.id); }}><DeleteIcon fontSize="inherit" /></IconButton>
                    </Box>
                  </ListItem>
                ))
              ) : (
                <Typography variant="body2" sx={{ fontSize: 16, mt: 1 }}>Select an invoice to see notes</Typography>
              )}
            </List>
          </CardContent>
        </Card>
      </Box>

      {/* Edit dialog for single or bulk edits */}
      <Dialog open={editDialogOpen} onClose={() => setEditDialogOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Edit selected invoice(s)</DialogTitle>
        <DialogContent>
          <Box sx={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 1, mt: 1 }}>
            <TextField
              label="Hub"
              size="small"
              value={editForm.hub}
              onChange={(e) => setEditForm(prev => ({ ...prev, hub: e.target.value }))}
            />
            <Select size="small" value={editForm.folder} onChange={(e) => setEditForm(prev => ({ ...prev, folder: e.target.value }))}>
              {FOLDERS.map(f => <MenuItem key={f} value={f}>{f}</MenuItem>)}
            </Select>
            <Select size="small" value={editForm.assigned} onChange={(e) => setEditForm(prev => ({ ...prev, assigned: e.target.value }))}>
              {ASSIGNEES.map(a => <MenuItem key={a} value={a}>{a}</MenuItem>)}
            </Select>
            <TextField size="small" label="Ref" value={editForm.ref} onChange={(e) => setEditForm(prev => ({ ...prev, ref: e.target.value.replace(/\D/g,'').slice(0,6) }))} />
            <Select size="small" value={editForm.status} onChange={(e) => setEditForm(prev => ({ ...prev, status: e.target.value }))}>
              {STATUS_OPTIONS.map(s => <MenuItem key={s} value={s}>{s}</MenuItem>)}
            </Select>
          </Box>
          <Typography variant="caption" color="text.secondary" sx={{ mt: 1, display: 'block' }}>
            Note: Only status, folder, assigned and ref are persisted to the server. Hub edits are local/UI-only unless server support is added.
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setEditDialogOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={async () => {
            const selected = getSelectedRows();
            if (!selected.length) return;

            // Build patch object with only allowed fields
            const allowedPatch = {};
            if (editForm.status) allowedPatch.status = editForm.status;
            if (editForm.folder) allowedPatch.folder = editForm.folder;
            if (editForm.assigned) allowedPatch.assigned = editForm.assigned;
            if (editForm.ref && /^\d{6}$/.test(editForm.ref)) allowedPatch.ref = editForm.ref;

            // Optimistic UI update
            setRows(prev => prev.map(r => selected.some(s => s.id === r.id) ? { ...r, ...editForm, ...allowedPatch } : r));

            // Send PATCH requests for each selected row, limited to allowed fields
            try {
              await Promise.all(selected.map(s => fetch(`${API_BASE}/api/invoices/${s.id}`, {
                method: 'PATCH',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(allowedPatch)
              })));
            } catch (e) {
              console.error('Bulk edit failed', e);
            }

            setEditDialogOpen(false);
          }}>Apply</Button>
        </DialogActions>
      </Dialog>

      {/* Notes editor */}
      <Dialog open={noteDialogOpen} onClose={() => setNoteDialogOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>
          {selectedNote ? (noteReadOnly ? "View Note" : "Edit Note") : "Add Note"}
          <IconButton
            aria-label="close"
            onClick={() => setNoteDialogOpen(false)}
            sx={{ position: "absolute", right: 8, top: 8 }}
          >
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        <DialogContent>
          {!selectedRow ? (
            <Typography variant="body2" color="text.secondary">
              Select an invoice to add a note.
            </Typography>
          ) : (
            <TextField
              autoFocus
              fullWidth
              multiline
              minRows={3}
              InputProps={{ readOnly: Boolean(noteReadOnly) }}
              placeholder="Type your note..."
              value={noteText}
              onChange={(e) => setNoteText(e.target.value)}
              sx={{ mt: 1 }}
            />
          )}
        </DialogContent>
        <DialogActions>
          {selectedNote && !noteReadOnly && (
            <Button
              color="error"
              onClick={() => {
                handleDeleteNote(selectedNote.id);
                setNoteDialogOpen(false);
              }}
            >
              Delete
            </Button>
          )}
          <Box sx={{ flex: 1 }} />
          <Button onClick={() => setNoteDialogOpen(false)}>Cancel</Button>
          {!noteReadOnly && (
            <Button
              variant="contained"
              onClick={handleSaveNote}
              disabled={!selectedRow || !noteText.trim()}
            >
              Save
            </Button>
          )}
        </DialogActions>
      </Dialog>

      {/* Match -> folder */}
      <Dialog open={folderDialogOpen} onClose={() => setFolderDialogOpen(false)} fullWidth>
        <DialogTitle>
          Select folder for Match
          <IconButton
            aria-label="close"
            onClick={() => setFolderDialogOpen(false)}
            sx={{ position: "absolute", right: 8, top: 8 }}
          >
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        <DialogContent>
          <Select
            fullWidth
            size="small"
            value={chosenFolder}
            onChange={(e) => setChosenFolder(e.target.value)}
            sx={{ mt: 1 }}
          >
            {FOLDERS.map((f) => (
              <MenuItem key={f} value={f}>{f}</MenuItem>
            ))}
          </Select>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setFolderDialogOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={confirmMatch}>Confirm</Button>
        </DialogActions>
      </Dialog>

      {/* Post -> assignee */}
      <Dialog open={assigneeDialogOpen} onClose={() => setAssigneeDialogOpen(false)} fullWidth>
        <DialogTitle>
          Select assignee for Post
          <IconButton
            aria-label="close"
            onClick={() => setAssigneeDialogOpen(false)}
            sx={{ position: "absolute", right: 8, top: 8 }}
          >
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        <DialogContent>
          <Select
            fullWidth
            size="small"
            value={chosenAssignee}
            onChange={(e) => setChosenAssignee(e.target.value)}
            sx={{ mt: 1 }}
          >
            {ASSIGNEES.map((p) => (
              <MenuItem key={p} value={p}>{p}</MenuItem>
            ))}
          </Select>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setAssigneeDialogOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={confirmPost}>Confirm</Button>
        </DialogActions>
      </Dialog>

      {/* Complete -> ref */}
      <Dialog open={refDialogOpen} onClose={() => setRefDialogOpen(false)} fullWidth>
        <DialogTitle>
          Enter 6-digit reference
          <IconButton
            aria-label="close"
            onClick={() => setRefDialogOpen(false)}
            sx={{ position: "absolute", right: 8, top: 8 }}
          >
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        <DialogContent>
          <TextField
            fullWidth
            inputMode="numeric"
            placeholder="e.g. 123456"
            value={refValue}
            onChange={(e) => setRefValue(e.target.value.replace(/\D/g, "").slice(0, 6))}
            sx={{ mt: 1 }}
          />
          <Typography
            variant="caption"
            color={/^\d{6}$/.test(refValue) ? "success.main" : "error"}
          >
            {/^(\d{6})$/.test(refValue) ? "Looks good." : "Reference must be exactly 6 digits."}
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setRefDialogOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={confirmComplete} disabled={!/^\d{6}$/.test(refValue)}>
            Confirm
          </Button>
        </DialogActions>
      </Dialog>

      {/* Delete confirmation */}
      <Dialog open={deleteDialog.open} onClose={cancelDeleteSelected}>
        <DialogTitle>
          Delete {deleteDialog.ids.length > 1 ? 'invoices' : 'invoice'}
        </DialogTitle>
        <DialogContent>
          <Typography variant="body2" sx={{ mt: 0.5 }}>
            {deleteDialog.message}
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={cancelDeleteSelected}>Cancel</Button>
          <Button
            onClick={confirmDeleteSelected}
            variant="contained"
            color="error"
          >
            Delete
          </Button>
        </DialogActions>
      </Dialog>

      {/* NEW: Outlook-like email compose */}
      <Dialog open={emailDialogOpen} onClose={() => setEmailDialogOpen(false)} fullWidth maxWidth="md">
        <DialogTitle sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <Box sx={{ display: 'flex', flexDirection: 'column' }}>
            <Typography variant="subtitle1" fontWeight={600}>New message</Typography>
            <Typography variant="caption" color="text.secondary">From shared mailbox</Typography>
          </Box>
          <Box sx={{ flex: 1 }} />
          <IconButton aria-label="close" onClick={() => setEmailDialogOpen(false)}>
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        <DialogContent dividers>
          <Box sx={{ display: 'grid', gridTemplateColumns: '90px 1fr', rowGap: 1.2, columnGap: 1.2 }}>
            <Typography variant="body2" sx={{ alignSelf: 'center' }}>From</Typography>
            <Select size="small" value={emailFrom} onChange={(e) => setEmailFrom(e.target.value)}>
              {SHARED_MAILBOXES.map((m) => (
                <MenuItem key={m} value={m}>{m}</MenuItem>
              ))}
            </Select>

            <Typography variant="body2" sx={{ alignSelf: 'center' }}>To</Typography>
            <Autocomplete
              freeSolo
              multiple
              options={peopleOptions}
              getOptionLabel={(opt) => (typeof opt === 'string' ? opt : `${opt.name} <${opt.email}>`) }
              filterOptions={(x) => x}
              value={parseEmailList(emailTo)}
              onChange={(_e, value) => {
                // value can be array of strings or {name,email}
                setEmailTo(value.map(v => (typeof v === 'string' ? v : v.email || v)).join('; '));
              }}
              onInputChange={(_e, input, _reason) => {
                const term = (input || '').trim();

                // clear any pending timer
                if (peopleFetchRef.current.timer) {
                  clearTimeout(peopleFetchRef.current.timer);
                  peopleFetchRef.current.timer = null;
                }

                // If input is empty, clear options and abort pending requests
                if (!term) {
                  setPeopleOptions([]);
                  peopleFetchRef.current.lastTerm = '';
                  if (peopleFetchRef.current.abort) {
                    try { peopleFetchRef.current.abort.abort(); } catch {};
                    peopleFetchRef.current.abort = null;
                  }
                  return;
                }

                // debounce network calls by 300ms
                peopleFetchRef.current.timer = setTimeout(async () => {
                  try {
                    // allow single-letter searches (term.length >= 1)
                    if (peopleFetchRef.current.lastTerm === term) return;
                    peopleFetchRef.current.lastTerm = term;

                    if (peopleFetchRef.current.abort) {
                      try { peopleFetchRef.current.abort.abort(); } catch {};
                    }
                    const controller = new AbortController();
                    peopleFetchRef.current.abort = controller;

                    const resp = await fetch(`${API_BASE}/api/graph/people?q=${encodeURIComponent(term)}`, { signal: controller.signal });
                    if (!resp.ok) {
                      const txt = await resp.text();
                      throw new Error(txt || 'People search failed');
                    }
                    const json = await resp.json();
                    setPeopleOptions(json || []);
                    // DEBUG: show fetched results in console to inspect
                    try { console.debug('peopleOptions set ->', json); } catch {}
                  } catch (err) {
                    if (err.name === 'AbortError') return;
                    setPeopleSearchError(err?.message || 'People search failed');
                  }
                }, 300);
              }}
              renderInput={(params) => (
                <TextField
                  {...params}
                  size="small"
                  placeholder="recipient@example.com; another@domain.com or search by name"
                  InputProps={{
                    ...params.InputProps,
                    endAdornment: (
                      <InputAdornment position="end">
                        <Typography variant="caption" color="text.secondary">semicolon separated</Typography>
                      </InputAdornment>
                    ),
                  }}
                />
              )}
              sx={{ minWidth: 320 }}
            />

            <Typography variant="body2" sx={{ alignSelf: 'center' }}>Cc</Typography>
            <TextField size="small" placeholder="optional" value={emailCc} onChange={(e) => setEmailCc(e.target.value)} />

            <Typography variant="body2" sx={{ alignSelf: 'center' }}>Bcc</Typography>
            <TextField size="small" placeholder="optional" value={emailBcc} onChange={(e) => setEmailBcc(e.target.value)} />

            <Typography variant="body2" sx={{ alignSelf: 'center' }}>Subject</Typography>
            <TextField size="small" placeholder="Subject" value={emailSubject} onChange={(e) => setEmailSubject(e.target.value)} />
          </Box>

          <Divider sx={{ my: 1.5 }} />

          {/* Attachments row */}
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
            <Button size="small" component="label" startIcon={<AttachFileIcon />}>
              Add attachments
              <input hidden type="file" multiple onChange={(e) => addFiles(e.target.files)} />
            </Button>
            <Button size="small" onClick={attachInvoicePdf} disabled={!selectedRow}>
              Attach selected invoice PDF
            </Button>
          </Box>

          {/* Attachment chips */}
          <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mb: 2 }}>
            {emailFiles.map((f, i) => (
              <Chip key={`${f.name}-${i}`} label={`${f.name} (${Math.round(f.size/1024)} KB)`} onDelete={() => removeFile(i)} />
            ))}
          </Box>

          {/* Body area: Outlook-like plain editor */}
          <TextField
            placeholder="Type your message…"
            multiline
            minRows={10}
            fullWidth
            value={emailBody}
            onChange={(e) => setEmailBody(e.target.value)}
          />
        </DialogContent>
        <DialogActions>
          <Box sx={{ flex: 1 }} />
          <Button onClick={() => setEmailDialogOpen(false)}>Discard</Button>
          <Button variant="contained" startIcon={<SendIcon />} onClick={handleSendEmail} disabled={isSending}>
            {isSending ? 'Sending…' : 'Send'}
          </Button>
        </DialogActions>
      </Dialog>

      {/* Snackbar */}
      <Snackbar
        open={snackbarOpen}
        anchorOrigin={{ vertical: "bottom", horizontal: "right" }}
        autoHideDuration={4000}
        onClose={() => setSnackbarOpen(false)}
        TransitionComponent={SlideUpTransition}
        message={`${newItemsCount} new item(s) added`}
        ContentProps={{
          sx: { backgroundColor: "#1976d2", color: "#fff", fontSize: 16, width: 250, maxWidth: "100%" },
        }}
      />
      <Snackbar
        open={Boolean(peopleSearchError)}
        autoHideDuration={5000}
        onClose={() => setPeopleSearchError(null)}
        message={peopleSearchError || ''}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
      />
    </Box>
  );
}

export default InvoiceDashboard;

/** Helper to safely map selected ids to rows (works with filtered / full data) */
function useCallbackIdsToRows(rowSelectionModel, rows) {
  const refRows = useRef(rows);
  useEffect(() => { refRows.current = rows; }, [rows]);

  return () => {
    const setIds = new Set(rowSelectionModel);
    return refRows.current.filter(r => setIds.has(r.id));
  };
}

/**
 * Parse a semicolon-separated email list into array for Autocomplete value.
 * Accepts strings like "Alice <a@x.com>; bob@x.com" or plain emails.
 */
function parseEmailList(s) {
  if (!s) return [];
  return String(s)
    .split(/;+/)
    .map(x => x.trim())
    .filter(Boolean)
    .map(item => {
      const m = item.match(/^(.*)\s*<([^>]+)>$/);
      if (m) return { name: m[1].trim(), email: m[2].trim() };
      return item;
    });
}
