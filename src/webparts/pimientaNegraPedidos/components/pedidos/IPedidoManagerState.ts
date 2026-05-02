import type { IPedido, IPedidoFormData } from '../../models/IPedido';
import type { IPedidoItemFormData } from '../../models/IPedidoItem';

export interface IPedidoManagerState {
  pedidos: IPedido[];
  loading: boolean;
  saving: boolean;
  error: string;
  success: string;
  isPanelOpen: boolean;
  editingPedido?: IPedido;
  formData: IPedidoFormData;
  itemsForm: IPedidoItemFormData[];
  loadingItems: boolean;
}
