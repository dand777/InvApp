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
} from "@mui/material";
import { DataGrid, GridToolbarQuickFilter } from "@mui/x-data-grid";
import EditIcon from "@mui/icons-material/Edit";
import DeleteIcon from "@mui/icons-material/Delete";
import CloseIcon from "@mui/icons-material/Close";

// Use env-driven API base; falls back to localhost for dev
const API_BASE = import.meta.env.VITE_API_URL || "http://localhost:5000";

const STATUS_OPTIONS = ["New", "Matched", "Posting", "Completed"];

const formatUKDate = (isoString) => {
  if (!isoString) return "";
  const date = new Date(isoString);
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
};

function QuickSearchToolbar() {
  return (
    <Box sx={{ p: 0.5, pb: 0 }}>
      <GridToolbarQuickFilter />
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

  // New: status filter (header dropdown)
  const [statusFilter, setStatusFilter] = useState("All");

  // --- New items pop-up state ---
  const [newItemsCount, setNewItemsCount] = useState(0);
  const [snackbarOpen, setSnackbarOpen] = useState(false);

  const prevIdsRef = useRef(new Set());

  // ------------------- Polling for live updates -------------------
  useEffect(() => {
    const fetchInvoices = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/invoices`);
        const data = await res.json();

        const prevIds = prevIdsRef.current;
        const newRows = data.map((row) => {
          if (!prevIds.has(row.id)) {
            return { ...row, isNew: true };
          }
          return { ...row, isNew: false };
        });

        // Count new items
        const newCount = newRows.filter((r) => r.isNew).length;
        if (newCount > 0) {
          setNewItemsCount(newCount);
          setSnackbarOpen(true);
        }

        prevIdsRef.current = new Set(data.map((r) => r.id));
        setRows(newRows);

        // Remove highlight after 4 seconds
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

  // ------------------- Keep selectedRow in sync with rows -----------
  const filteredRows = useMemo(() => {
    let base = filterSupplier ? rows.filter((r) => r.supplier === filterSupplier) : rows;

    // Apply Status filter (treat missing as 'New')
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

  // ------------------- Notes Handlers -------------------
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
      // UPDATE existing
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
      // CREATE new
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
      await fetch(`${API_BASE}/api/invoices/${invoiceId}/notes/${noteId}`, {
        method: "DELETE",
      });

      const updatedRow = {
        ...selectedRow,
        notes: (Array.isArray(selectedRow.notes) ? selectedRow.notes : []).filter(
          (n) => n.id !== noteId
        ),
      };

      setRows((prev) => prev.map((r) => (r.id === invoiceId ? updatedRow : r)));
      setSelectedRow(updatedRow);
    } catch (err) {
      console.error(err);
    }
  };

  // ------------------- Suppliers -------------------
  const supplierCounts = useMemo(() => {
    return rows.reduce((acc, row) => {
      acc[row.supplier] = (acc[row.supplier] || 0) + 1;
      return acc;
    }, {});
  }, [rows]);

  const handleSupplierClick = (supplier) => {
    setFilterSupplier(supplier);
  };

  // ------------------- Columns (with Status + header filter) --------
  const columns = useMemo(
    () => [
      {
        field: "supplier",
        headerName: "Supplier",
        flex: 1,
      },
      { field: "hub", headerName: "Hub", flex: 1 },
      { field: "type", headerName: "Type", flex: 1 },
      {
        field: "status",
        headerName: "Status",
        flex: 1,
        // Display 'New' if missing
        valueGetter: (params) => params.row.status || "New",
        // Header with dropdown filter
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
    ],
    [statusFilter]
  );

  return (
    <Box sx={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <Typography variant="h5" gutterBottom>
        Invoice Manager
      </Typography>

      {/* DataGrid */}
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
          onRowSelectionModelChange={(newModel) => {
            setRowSelectionModel(newModel);
          }}
          // Clicking supplier cell should select row AND tick the checkbox
          onCellClick={(params) => {
            if (params.field === "supplier") {
              setRowSelectionModel([params.id]);
              setSelectedRow(params.row);
            }
          }}
          slots={{ toolbar: QuickSearchToolbar }}
          sx={{
            fontSize: 13,
            "& .MuiDataGrid-columnHeaders": {
              fontSize: 12,
              color: "#68adf1ff",
            },
            "& .MuiDataGrid-cell": {
              py: 0,
              lineHeight: "14px",
              display: "flex",
              alignItems: "center",
            },
            "& .MuiDataGrid-row.Mui-selected, .MuiDataGrid-row.Mui-selected:hover": {
              backgroundColor: "rgba(25,118,210,0.2)",
            },
            "& .MuiCheckbox-root": {
              width: "10px",
              height: "10px",
              padding: 0,
            },
            "& .MuiCheckbox-root svg": {
              transform: "scale(0.5)",
              transformOrigin: "center",
            },
          }}
        />
      </Box>

      <Box sx={{ display: "flex", mt: 1, flexShrink: 0 }}>
        {/* Suppliers Panel */}
        <Card sx={{ flex: 1, mr: 1, height: "320px", position: "relative" }}>
          {/* Sticky header */}
          <Box
            display="flex"
            justifyContent="space-between"
            alignItems="center"
            sx={{
              position: "sticky",
              top: 0,
              background: "transparent",
              zIndex: 1,
              p: 1,
              pb: 0.5,
            }}
          >
            <Typography variant="body2" sx={{ fontSize: 14 }}>
              Suppliers
            </Typography>
            <Button
              size="small"
              sx={{ fontSize: 14, textTransform: "none" }}
              onClick={() => setFilterSupplier(null)}
            >
              Show All
            </Button>
          </Box>

          {/* Scrollable list */}
          <Box sx={{ overflowY: "auto", height: "calc(100% - 32px)", px: 1 }}>
            <List dense>
              {Object.entries(supplierCounts).map(([supplier, count]) => (
                <ListItem
                  key={supplier}
                  button
                  onClick={() => handleSupplierClick(supplier)}
                  sx={{ py: 0.3 }}
                >
                  <ListItemText
                    primaryTypographyProps={{ fontSize: 10, color: "inherit" }}
                    primary={`${supplier} (${count})`}
                  />
                </ListItem>
              ))}
            </List>
          </Box>
        </Card>

        {/* Notes Panel */}
        <Card sx={{ flex: 2, height: "320px" }}>
          <CardContent sx={{ p: 1, height: "100%", overflowY: "auto" }}>
            <Box
              display="flex"
              justifyContent="space-between"
              alignItems="center"
              mb={0.5}
            >
              <Typography variant="body2" sx={{ fontSize: 14 }}>
                Notes
              </Typography>
              <Button
                size="small"
                sx={{ fontSize: 14, textTransform: "none" }}
                onClick={() => {
                  setSelectedNote(null);
                  setNoteText("");
                  setNoteDialogOpen(true);
                }}
              >
                + Add Note
              </Button>
            </Box>
            <List dense>
              {selectedRow && Array.isArray(selectedRow.notes) && selectedRow.notes.length > 0 ? (
                selectedRow.notes.map((note) => (
                  <ListItem
                    key={note.id}
                    sx={{
                      py: 0.3,
                      display: "flex",
                      justifyContent: "space-between",
                    }}
                  >
                    <ListItemText
                      primaryTypographyProps={{ fontSize: 10 }}
                      primary={`${note.text.substring(0, 20)}... (${note.date})`}
                      onClick={() => handleOpenNote(note)}
                    />
                    <Box>
                      <IconButton size="small" onClick={() => handleOpenNote(note)}>
                        <EditIcon fontSize="inherit" />
                      </IconButton>
                      <IconButton
                        size="small"
                        onClick={() => handleDeleteNote(note.id)}
                      >
                        <DeleteIcon fontSize="inherit" />
                      </IconButton>
                    </Box>
                  </ListItem>
                ))
              ) : (
                <Typography variant="body2" sx={{ fontSize: 10, mt: 1 }}>
                  Select an invoice to see notes
                </Typography>
              )}
            </List>
          </CardContent>
        </Card>
      </Box>

      {/* Note Dialog */}
      <Dialog open={noteDialogOpen} onClose={() => setNoteDialogOpen(false)} fullWidth>
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
          <TextField
            fullWidth
            multiline
            rows={4}
            value={noteText}
            onChange={(e) => setNoteText(e.target.value)}
            sx={{ fontSize: 12 }}
          />
        </DialogContent>
        <DialogActions>
          <Button onClick={handleSaveNote} variant="contained" size="small">
            Save
          </Button>
        </DialogActions>
      </Dialog>

      {/* New Items Pop-up */}
      <Snackbar
        open={snackbarOpen}
        anchorOrigin={{ vertical: "bottom", horizontal: "right" }}
        autoHideDuration={4000}
        onClose={() => setSnackbarOpen(false)}
        message={`${newItemsCount} new item(s) added`}
        ContentProps={{
          sx: {
            backgroundColor: "#1976d2",
            color: "#fff",
            fontSize: 10,
            width: 250,
            maxWidth: "100%",
          },
        }}
      />

      <style>{`
        .row-new-highlight {
          background-color: rgba(144, 238, 144, 0.5);
          transition: background-color 0.5s ease;
        }

        /* Uniform scrollbars matching DataGrid background */
        * {
          scrollbar-width: thin;
          scrollbar-color: rgba(39, 52, 167, 1) #191d42ff;
        }
        ::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }
        ::-webkit-scrollbar-track {
          background: #f5f5f5;
          border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb {
          background-color: #f5f5f5;
          border-radius: 4px;
          border: 2px solid #f5f5f5;
        }
        ::-webkit-scrollbar-thumb:hover {
          background-color: #e0e0e0;
        }
      `}</style>
    </Box>
  );
}

export default InvoiceDashboard;
