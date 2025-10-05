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
  Slide,
  Chip,
  Tooltip,
  Divider,
  InputAdornment,
} from "@mui/material";
import { DataGrid, GridToolbarQuickFilter } from "@mui/x-data-grid";
import EditIcon from "@mui/icons-material/Edit";
import DeleteIcon from "@mui/icons-material/Delete";
import CloseIcon from "@mui/icons-material/Close";
import SendIcon from "@mui/icons-material/Send";
import AttachFileIcon from "@mui/icons-material/AttachFile";

// same-origin in Azure; optional override for local dev
const API_BASE = (import.meta.env.VITE_API_URL ?? '').trim().replace(/\/+$/, '');


const STATUS_OPTIONS = ["New", "Matched", "Posting", "Completed"];
const FOLDERS = ["UK", "Ireland", "Foreign", "Spain", "GmbH"];
const ASSIGNEES = ["Allan Perry", "Marcos Silva", "Caroline Stathatos"];

// Optional: one or more shared mailboxes to send from
const SHARED_MAILBOXES = (
  import.meta.env.VITE_SHARED_MAILBOXES?.split(",") || [
    "ap@gear4music.com",
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

function SlideUpTransition(props) {
  return <Slide {...props} direction="up" />;
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
      </Box>
    </Box>
  );
}

function InvoiceDashboard() {
  const [rows, setRows] = useState([]);
  const [filterSupplier, setFilterSupplier] = useState(null);

  const [selectedRow, setSelectedRow] = useState(null);
  const [rowSelectionModel, setRowSelectionModel] = useState([]);

  const [selectedNote, setSelectedNote] = useState(null);
  const [noteDialogOpen, setNoteDialogOpen] = useState(false);
  const [noteText, setNoteText] = useState("");

  // Status column header filter
  const [statusFilter, setStatusFilter] = useState("All");

  // New items pop-up
  const [newItemsCount, setNewItemsCount] = useState(0);
  const [snackbarOpen, setSnackbarOpen] = useState(false);

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

  const supplierCounts = useMemo(() => {
    return rows.reduce((acc, row) => {
      acc[row.supplier] = (acc[row.supplier] || 0) + 1;
      return acc;
    }, {});
  }, [rows]);

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
    let base = filterSupplier ? rows.filter((r) => r.supplier === filterSupplier) : rows;
    if (statusFilter !== "All") {
      base = base.filter((r) => (r.status || "New") === statusFilter);
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
  const handleOpenNote = (note) => {
    if (!selectedRow) return;
    setSelectedNote(note);
    setNoteText(note.text);
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
    // No-op for now — but wired up for future.
    console.log("Edit clicked for rows:", selected.map(r => r.id));
  };

  const handleDeleteSelected = async () => {
    const selected = getSelectedRows();
    if (selected.length === 0) return;

    const msg =
      selected.length === 1
        ? `Delete invoice ${selected[0].id}?`
        : `Delete ${selected.length} selected invoices?`;
    if (!window.confirm(msg)) return;

    // Optimistic removal
    const ids = selected.map(r => r.id);
    setRows(prev => prev.filter(r => !ids.includes(r.id)));
    // Clear selection after delete
    setRowSelectionModel([]);

    // Try API deletes in the background — if unsupported, UI still updates.
    try {
      await Promise.allSettled(
        ids.map(id =>
          fetch(`${API_BASE}/api/invoices/${id}`, { method: "DELETE" })
        )
      );
    } catch (e) {
      console.error("Delete error:", e);
      // (Optional) You could refetch here if you want to re-sync.
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

      // Backend should send via Microsoft Graph using the shared mailbox and return a JSON payload
      // e.g. { id: "<messageId>", bodySaved: "<string to add to notes>" }
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
      { field: "supplier", headerName: "Supplier", flex: 1 },
      { field: "hub", headerName: "Hub", flex: 1 },
      { field: "type", headerName: "Type", flex: 1 },
      { field: "invoiceno", headerName: "Invoice No", flex: 1 },

      {
        field: "invoice_date",
        headerName: "Invoice Date",
        flex: 1,
        valueGetter: (params) => formatUKDate(params.row.invoice_date),
      },
      { field: "po", headerName: "PO", flex: 1 },
      { field: "folder", headerName: "Folder", flex: 1 },
      { field: "assigned", headerName: "Assigned", flex: 1 },
      { field: "ref", headerName: "Ref", flex: 1 },
      {
        field: "last_modified",
        headerName: "Last Modified",
        flex: 1,
        valueGetter: (params) => formatUKDate(params.row.last_modified),
      },
      {
        field: "created_on",
        headerName: "Created On",
        flex: 1,
        valueGetter: (params) => formatUKDate(params.row.created_on),
      },
      {
        field: "status",
        headerName: "Status",
        flex: 1,
        valueGetter: (params) => params.row.status || "New",
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
    ],
    [statusFilter]
  );

  return (
    <Box sx={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <Typography variant="h5" gutterBottom>
        Invoice Manager
      </Typography>

      {/* DataGrid with custom toolbar */}
      <Box sx={{ flexGrow: 1, minHeight: 0 }}>
        <DataGrid
          rows={filteredRows}
          columns={columns}
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
            <Button size="small" sx={{ fontSize: 14, textTransform: "none" }} onClick={() => setFilterSupplier(null)}>
              Show All
            </Button>
          </Box>
          <Box sx={{ overflowY: "auto", height: "calc(100% - 32px)", px: 1 }}>
            <List dense>
              {Object.entries(supplierCounts).map(([supplier, count]) => (
                <ListItem key={supplier} button onClick={() => setFilterSupplier(supplier)} sx={{ py: 0.3 }}>
                  <ListItemText
                    primaryTypographyProps={{ fontSize: 14, color: "inherit" }}
                    primary={`${supplier} (${count})`}
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
                  onClick={() => { setSelectedNote(null); setNoteText(""); setNoteDialogOpen(true); }}
                  disabled={!selectedRow}
                >
                  + Add Note
                </Button>
                {/* NEW: Email button next to + Add Note */}
                <Button
                  size="small"
                  variant="contained"
                  startIcon={<SendIcon fontSize="small" />}
                  sx={{ fontSize: 14, textTransform: "none" }}
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
                  <ListItem key={note.id} sx={{ py: 0.3, display: "flex", justifyContent: "space-between" }}>
                    <ListItemText
                      primaryTypographyProps={{ fontSize: 14 }}
                      primary={`${note.text.substring(0, 20)}... (${note.date})`}
                      onClick={() => handleOpenNote(note)}
                    />
                    <Box>
                      <IconButton size="small" onClick={() => handleOpenNote(note)}><EditIcon fontSize="inherit" /></IconButton>
                      <IconButton size="small" onClick={() => handleDeleteNote(note.id)}><DeleteIcon fontSize="inherit" /></IconButton>
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

      {/* Notes editor */}
      <Dialog open={noteDialogOpen} onClose={() => setNoteDialogOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>
          {selectedNote ? "Edit Note" : "Add Note"}
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
              placeholder="Type your note..."
              value={noteText}
              onChange={(e) => setNoteText(e.target.value)}
              sx={{ mt: 1 }}
            />
          )}
        </DialogContent>
        <DialogActions>
          {selectedNote && (
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
          <Button
            variant="contained"
            onClick={handleSaveNote}
            disabled={!selectedRow || !noteText.trim()}
          >
            Save
          </Button>
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

      {/* NEW: Outlook-like email compose */}
      <Dialog open={emailDialogOpen} onClose={() => setEmailDialogOpen(false)} fullWidth maxWidth="md">
        <DialogTitle sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <Box sx={{ display: 'flex', flexDirection: 'column' }}>
            <Typography variant="subtitle1" fontWeight={600}>New message</Typography>
            <Typography variant="caption" color="text.secondary">From shared mailbox</Typography>
          </Box>
          <Box sx={{ flex: 1 }} />
          <Tooltip title="Send">
            <span>
              <Button
                onClick={handleSendEmail}
                startIcon={<SendIcon />}
                variant="contained"
                disabled={isSending}
              >
                {isSending ? 'Sending…' : 'Send'}
              </Button>
            </span>
          </Tooltip>
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
            <TextField
              size="small"
              placeholder="recipient@example.com; another@domain.com"
              value={emailTo}
              onChange={(e) => setEmailTo(e.target.value)}
              InputProps={{
                endAdornment: (
                  <InputAdornment position="end">
                    <Typography variant="caption" color="text.secondary">semicolon separated</Typography>
                  </InputAdornment>
                ),
              }}
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
          sx: { backgroundColor: "#1976d2", color: "#fff", fontSize: 10, width: 250, maxWidth: "100%" },
        }}
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