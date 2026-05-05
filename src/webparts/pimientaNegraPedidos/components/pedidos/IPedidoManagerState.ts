import type { IPedido, IPedidoFormData } from '../../models/IPedido';
import type { IPedidoItem, IPedidoItemFormData } from '../../models/IPedidoItem';

export type PedidoView = 'lista' | 'detalle';

export interface IPedidoFiltros {
  texto: string;
  estado: string;
  franja: string;
  metodo: string;
}

export interface IPedidoManagerState {
  // Vista
  view: PedidoView;
  viewMode: 'tarjetas' | 'lista';
  showDetailPanel: boolean;
  selectedPedido?: IPedido;

  // Lista
  pedidos: IPedido[];
  loading: boolean;
  error: string;
  success: string;
  filtros: IPedidoFiltros;
  pedidosSeleccionados: number[];
  soloHoy: boolean;

  // Detalle
  pedidoItems: IPedidoItem[];
  loadingItems: boolean;
  costoEnvio: number;
  savingEstado: boolean;

  // ABM
  isPanelOpen: boolean;
  saving: boolean;
  editingPedido?: IPedido;
  formData: IPedidoFormData;
  itemsForm: IPedidoItemFormData[];
}
