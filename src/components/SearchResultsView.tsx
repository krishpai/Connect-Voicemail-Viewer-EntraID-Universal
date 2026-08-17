import React, { useMemo, useState, useCallback, useRef } from "react";
import { apiRequest } from "../authConfig";

import {
  DataGridPro,
  GridFooterContainer,
  GridColumnMenu,
  type GridColDef,
  type GridRowSelectionModel,
  type GridColumnMenuProps,
  type GridColumnVisibilityModel,
} from '@mui/x-data-grid-pro';

import {
  Tooltip,
  IconButton,
  Dialog,
  DialogActions,
  DialogContent,
  DialogContentText,
  DialogTitle,
  Button,
  CircularProgress,
  Box,
  TablePagination,
} from '@mui/material';

import MailOutlineIcon from '@mui/icons-material/MailOutline';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import PhoneIcon from '@mui/icons-material/Phone';
import DeleteIcon from '@mui/icons-material/Delete';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';
import TranscriptPopup from './TranscriptPopup';
import { useAcquireTokenWithRecovery } from "../hooks/useAcquireTokenWithRecovery";

// --- API Endpoints ---
const API_ENDPOINT_ENTRA_AUTH = import.meta.env.VITE_API_URL_ENTRA_AUTH;
//const API_ENDPOINT_CONNECT_AUTH = import.meta.env.VITE_API_URL_CONNECT_AUTH;
const isIframe = window.self !== window.top;

// --- Interfaces ---
interface SearchResultsViewProps {
  searchResult: string | null;
  canDeleteVM: string | null | undefined;

  onDialNumberClicked: (value: string, contactid: string) => void;
}

interface MatchedObject {
  vmx3_unread?: string;
  vmx3_contact_id: string;
  vmx3_customer_number: string;
  vmx3_queue_name: string;
  vmx3_target: string;
  vmx3_preferred_agent: string;
  vmx3_region: string;
  vmx3_timestamp: string;
  vmx3_lang_value: string;
  vmx3_call_category: string;
  vmx3_dialed_number: string;
  vmx3_queue: string;
  transcript: string;
  presigned_url: string;
}

interface GridRow extends MatchedObject {
  id: string;
  fileName: string;
}

/**
 * CUSTOM FOOTER PROPS INTERFACE
 * We explicitly define the types for the props we pass through slotProps.
 */
interface CustomFooterProps extends React.HTMLAttributes<HTMLDivElement> {
  contactId?: string | null;
  count?: number;
  page?: number;
  pageSize?: number;
  onPageChange?: (event: React.MouseEvent<HTMLButtonElement> | null, newPage: number) => void;
  onPageSizeChange?: (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => void;
}

// --- Sub-Components ---

const CustomFooter = (props: CustomFooterProps) => {
  const {
    contactId,
    count = 0,
    page = 0,
    pageSize = 15,
    onPageChange,
    onPageSizeChange,
    ...other
  } = props;

  const [copied, setCopied] = useState(false);

  const handleCopyContactId = async () => {
    if (!contactId) return;
    try {
      await navigator.clipboard.writeText(contactId);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) { console.error("Copy failed", err); }
  };

  return (
    <GridFooterContainer {...other} sx={{
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      minHeight: '52px !important',
      paddingY: 0
    }}>
      <Box sx={{ pl: 2, display: 'flex', alignItems: 'center', gap: 1 }}>
        <Box sx={{ fontSize: '0.875rem', color: '#666', fontWeight: 500, lineHeight: 1 }}>
          {contactId ? `Selected Contact ID: ${contactId}` : ''}
        </Box>
        {contactId && (
          <Tooltip title={copied ? "Copied!" : "Copy Contact ID"}>
            <IconButton size="small" onClick={handleCopyContactId}>
              {copied ? <CheckCircleIcon fontSize="small" color="success" /> : <ContentCopyIcon fontSize="small" />}
            </IconButton>
          </Tooltip>
        )}
      </Box>

      <TablePagination
        component="div"
        count={count}
        page={page}
        onPageChange={onPageChange ?? (() => { })}
        rowsPerPage={pageSize}
        onRowsPerPageChange={onPageSizeChange}
        rowsPerPageOptions={[15, 25, 50]}
        sx={{
          border: 'none',
          '& .MuiTablePagination-toolbar': {
            minHeight: '52px',
            height: '52px',
            display: 'flex',
            alignItems: 'center',
            paddingY: 0,
          },
          '& .MuiTablePagination-selectLabel': {
            margin: 0,
            lineHeight: 1,
            display: 'flex',
            alignItems: 'center'
          },
          '& .MuiTablePagination-displayedRows': {
            margin: 0,
            lineHeight: 1,
            display: 'flex',
            alignItems: 'center'
          },
        }}
      />
    </GridFooterContainer>
  );
};

const CustomColumnMenu = (props: GridColumnMenuProps) => (
  <GridColumnMenu
    {...props}
    slots={{
      columnMenuHideColumnItem: null,
      columnMenuManageColumnsItem: null,
      columnMenuColumnsItem: null
    }}
  />
);

const NoRowsOverlay = () => (
  <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'gray' }}>
    No matching recordings found.
  </Box>
);

// --- Main Component ---

export const SearchResultsView: React.FC<SearchResultsViewProps> = ({ searchResult, canDeleteVM, onDialNumberClicked }) => {
  const acquireTokenWithRecovery = useAcquireTokenWithRecovery();
  const playingAudioRef = useRef<HTMLAudioElement | null>(null);

  const [readMessages, setReadMessages] = useState<Set<string>>(new Set());
  const [deletedFileNames, setDeletedFileNames] = useState<Set<string>>(new Set());
  const [columnVisibilityModel, setColumnVisibilityModel] = useState<GridColumnVisibilityModel>({ id: false });
  const [rowSelectionModel, setRowSelectionModel] = useState<GridRowSelectionModel>({ type: 'include', ids: new Set() });
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [itemToDelete, setItemToDelete] = useState<{ id: string, fileName: string } | null>(null);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [paginationModel, setPaginationModel] = useState({ pageSize: 15, page: 0 });

  const gridRows = useMemo<GridRow[]>(() => {
    if (!searchResult) return [];
    try {
      const data = JSON.parse(searchResult);
      const rawData: Record<string, MatchedObject> = data.matched_objects || {};
      return Object.entries(rawData)
        .filter(([fileName]) => !deletedFileNames.has(fileName))
        .map(([fileName, details]) => ({
          id: details.vmx3_contact_id,
          fileName,
          ...details,
          vmx3_unread: readMessages.has(details.vmx3_contact_id) ? 'N' : details.vmx3_unread,
          vmx3_queue_name: details.vmx3_queue_name
          //vmx3_queue_name: (details.vmx3_target === "agent" && details.vmx3_preferred_agent?.toLowerCase() === userName?.toLowerCase())
          //  ? "Self" : (details.vmx3_queue_name === 'VMX3_VM_QUEUE' ? 'Self' : details.vmx3_queue_name)
        }))

    } catch (e) { console.log(e); return []; }
  }, [searchResult, readMessages, deletedFileNames]);

  const selectedContactId = useMemo(() => {
    const selectionIds = rowSelectionModel.ids;
    if (!selectionIds || selectionIds.size === 0 || selectionIds.size > 1) return null;
    const firstId = selectionIds.values().next().value as string;
    return gridRows.find((row) => row.id === firstId)?.vmx3_contact_id || null;
  }, [rowSelectionModel, gridRows]);

  const handleAudioPlay = useCallback((e: React.SyntheticEvent<HTMLAudioElement>) => {
    if (playingAudioRef.current && playingAudioRef.current !== e.currentTarget) playingAudioRef.current.pause();
    playingAudioRef.current = e.currentTarget;
  }, []);

  const handleMarkAsRead = useCallback(async (contactId: string, fileName: string) => {
    setReadMessages(prev => new Set(prev).add(contactId));
    const apiUrl = `${API_ENDPOINT_ENTRA_AUTH}?function_code=mark_voice_message_read&vmx3_file_name=${fileName}`;
    try {
      // Fired by the audio element's onEnded, not by a click - there is no
      // user gesture to open a popup with, so stay silent. Marking as read is
      // best effort; the row is already shown as played in local state.
      const authResult = await acquireTokenWithRecovery({ ...apiRequest }, { allowInteraction: false });
      const token = authResult?.accessToken;
      if (!token) {
        console.warn("Skipping mark-as-read: no access token available.");
        return;
      }
      await fetch(apiUrl, { headers: { Authorization: `Bearer ${token}` } });
    } catch (error) { console.error("Mark read error:", error); }
  }, [acquireTokenWithRecovery]);

  const confirmDelete = useCallback(async () => {
    if (!itemToDelete) return;
    setIsDeleting(true);
    const apiUrl = `${API_ENDPOINT_ENTRA_AUTH}?function_code=delete_voice_message&vmx3_file_name=${itemToDelete.fileName}`;
    setDeleteError(null);
    try {
      // Runs from the Delete button, so the user gesture is still active and
      // the hook may open a popup if the token needs refreshing.
      const authResult = await acquireTokenWithRecovery({ ...apiRequest }, { allowInteraction: true });
      const token = authResult?.accessToken;
      if (!token) {
        setDeleteError("Sign in did not complete. Try again.");
        return;
      }

      const resp = await fetch(apiUrl, { headers: { Authorization: `Bearer ${token}` } });

      if (resp.status === 401 || resp.status === 403) {
        setDeleteError("Your session expired. Try again to sign in.");
        return;
      }

      if (resp.ok) {
        setDeletedFileNames(prev => new Set(prev).add(itemToDelete.fileName));
        setDeleteDialogOpen(false);
      } else {
        setDeleteError("The message could not be deleted. Try again.");
      }
    } catch (e) {
      console.error("Delete error:", e);
      setDeleteError("The message could not be deleted. Try again.");
    } finally { setIsDeleting(false); }
  }, [itemToDelete, acquireTokenWithRecovery]);

  const columns = useMemo<GridColDef<GridRow>[]>(() => {
    const baseColumns: GridColDef<GridRow>[] = [
      { field: 'id', filterable: false, headerName: 'Contact ID', width: 120, align: 'center', getApplyQuickFilterFn: () => null },
      {
        field: 'vmx3_unread', filterable: false, headerName: '', width: 70, align: 'center', getApplyQuickFilterFn: () => null, renderCell: (params) => (
          <Tooltip title={params.value === 'Y' ? "Unread" : "Played"}>
            <Box sx={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%' }}>
              {params.value === 'Y' ? <MailOutlineIcon color="primary" /> : <CheckCircleIcon color="action" />}
            </Box>
          </Tooltip>
        )
      },
      { field: 'vmx3_timestamp', headerName: 'Local Time', headerAlign: 'center', width: 220, align: 'center', valueFormatter: (value) => value ? new Date(value as string).toLocaleString() : '' },
      { field: 'vmx3_queue_name', headerName: 'Queue/Agent', headerAlign: 'center', width: 210, align: 'center' },
      { field: 'vmx3_customer_number', headerName: 'Caller number', headerAlign: 'center', width: 130, align: 'center' },
      { field: 'vmx3_dialed_number', headerName: 'Dialed number', headerAlign: 'center', width: 130, align: 'center' },
      { field: 'vmx3_lang_value', headerName: 'Language', headerAlign: 'center', width: 100, align: 'center' },
      {
        field: 'presigned_url', filterable: false, sortable: false, headerName: 'Listen', headerAlign: 'center', width: 220, align: 'center', renderCell: (params) => (
          <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
            <audio controls src={params.value as string} onPlay={handleAudioPlay} onEnded={() => handleMarkAsRead(params.row.id, params.row.fileName)} style={{ height: '24px', width: '250px' }} />
          </Box>
        )
      },
      {
        field: 'transcript', filterable: false, sortable: false, headerName: 'Transcript', headerAlign: 'center', width: 120, align: 'center', getApplyQuickFilterFn: () => null, renderCell: (params) => (
          <Box sx={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%' }}>
            <TranscriptPopup text={(params.value as string) ?? ""} />
          </Box>
        )
      },
      {
        field: 'dial_action', headerName: 'Call back', sortable: false, width: 90, align: 'center', getApplyQuickFilterFn: () => null, renderCell: (params) => (
          <IconButton color="primary" onClick={() => onDialNumberClicked(params.row.vmx3_customer_number, params.row.vmx3_contact_id)}><PhoneIcon /></IconButton>
        )
      },
      {
        field: 'delete_action', filterable: false, sortable: false, headerName: '', width: 90, align: 'center', getApplyQuickFilterFn: () => null, renderCell: (params) => (canDeleteVM === 'Y') ? (
          <IconButton onClick={() => { setItemToDelete({ id: params.row.id, fileName: params.row.fileName }); setDeleteError(null); setDeleteDialogOpen(true); }}><DeleteIcon /></IconButton>
        ) : null
      }
    ];
    return baseColumns.filter(col => isIframe || col.field !== 'dial_action');
  }, [canDeleteVM, onDialNumberClicked, handleAudioPlay, handleMarkAsRead]);

  if (!searchResult) return <Box sx={{ p: 5, textAlign: 'center' }}>No search performed.</Box>;

  return (
    <Box sx={{ height: 600, width: '100%', pt: 2 }}>
      <DataGridPro
        disableColumnMenu
        disableColumnSelector
        pagination
        showToolbar
        rows={gridRows}
        columns={columns}
        columnVisibilityModel={columnVisibilityModel}
        onColumnVisibilityModelChange={setColumnVisibilityModel}
        paginationModel={paginationModel}
        onPaginationModelChange={setPaginationModel}
        rowSelectionModel={rowSelectionModel}
        onRowSelectionModelChange={setRowSelectionModel}
        hideFooterSelectedRowCount
        slots={{
          columnMenu: CustomColumnMenu,
          footer: CustomFooter,
          noRowsOverlay: NoRowsOverlay,
        }}
        slotProps={{
          footer: {
            contactId: selectedContactId,
            count: gridRows.length,
            page: paginationModel.page,
            pageSize: paginationModel.pageSize,
            onPageChange: (_event: React.MouseEvent<HTMLButtonElement> | null, newPage: number) =>
              setPaginationModel(prev => ({ ...prev, page: newPage })),
            onPageSizeChange: (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) =>
              setPaginationModel(prev => ({ ...prev, pageSize: parseInt(event.target.value, 15), page: 0 })),
          } as CustomFooterProps, // Use our defined interface instead of 'any'
          toolbar: {
            showQuickFilter: true,

            printOptions: { disableToolbarButton: true },
            style: { backgroundColor: '#e0e0e0' },
          }
        }}
        sx={{
          '& .MuiDataGrid-columnHeader': { backgroundColor: '#2e2c2c33 !important', color: 'black !important' },
          '& .MuiDataGrid-columnHeaderTitle': { fontWeight: 'bold' },
          '& .MuiDataGrid-toolbarContainer': {
            display: 'flex',
            flexDirection: 'column',
            backgroundColor: '#424242 !important',
          },
        }}
      />

      <Dialog open={deleteDialogOpen} onClose={() => { setDeleteDialogOpen(false); setDeleteError(null); }}>
        <DialogTitle>Confirm Deletion</DialogTitle>
        <DialogContent>
          <DialogContentText>Permanently delete this voicemail?</DialogContentText>
          {deleteError && (
            <DialogContentText sx={{ mt: 2, color: 'error.main' }}>{deleteError}</DialogContentText>
          )}
        </DialogContent>
        <DialogActions sx={{ pb: 2, px: 3 }}>
          <Button onClick={() => { setDeleteDialogOpen(false); setDeleteError(null); }} color="inherit">Cancel</Button>
          <Button color="error" variant="contained" disabled={isDeleting} onClick={confirmDelete}>
            {isDeleting ? <CircularProgress size={24} color="inherit" /> : "Delete"}
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
};