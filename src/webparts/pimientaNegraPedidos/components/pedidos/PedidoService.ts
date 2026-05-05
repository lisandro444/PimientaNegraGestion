import type { IPedido, IPedidoFormData } from '../../models/IPedido';
import type { IPedidoItem, IPedidoItemFormData } from '../../models/IPedidoItem';
import { SharePointService } from '../../services/SharePointService';

type SharePointPedidoItem = {
  ID: number;
  Title?: string;
  FormResponseId?: number;
  SubmittedAt?: string;
  NombreCompleto?: string;
  WhatsApp?: string;
  MetodoEntrega?: string;
  DireccionEntrega?: string;
  FechaEntrega?: string;
  HorarioAproximado?: string;
  MetodoPago?: string;
  CubiertosDescartables?: boolean;
  Comentarios?: string;
  EstadoPedido?: string;
};

type SharePointPedidoItemRow = {
  ID: number;
  Title?: string;
  Pedido?: string | number;
  PedidoId?: number;
  PedidoFormResponseId?: number;
  Producto?: string;
  Cantidad?: number;
  Orden?: number;
};

const COSTO_ENVIO_KEY = 'pn_costo_envio';
const DEFAULT_COSTO_ENVIO = 2000;

export class PedidoService {
  private readonly sharePointService: SharePointService;

  private readonly pedidosListTitle: string = 'Pedidos';
  private readonly itemsListTitle: string = 'Productos';

  private readonly pedidoSelectFields: string[] = [
    'ID', 'Title', 'FormResponseId', 'SubmittedAt',
    'NombreCompleto', 'WhatsApp', 'MetodoEntrega', 'DireccionEntrega',
    'FechaEntrega', 'HorarioAproximado', 'MetodoPago', 'CubiertosDescartables',
    'Comentarios', 'EstadoPedido'
  ];

  private readonly itemSelectFields: string[] = [
    'ID', 'Title', 'PedidoFormResponseId', 'Producto', 'Cantidad'
  ];

  constructor(sharePointService: SharePointService) {
    this.sharePointService = sharePointService;
  }

  public getCostoEnvio(): number {
    if (typeof localStorage === 'undefined') return DEFAULT_COSTO_ENVIO;
    const stored = localStorage.getItem(COSTO_ENVIO_KEY);
    return stored ? parseFloat(stored) : DEFAULT_COSTO_ENVIO;
  }

  public setCostoEnvio(valor: number): void {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(COSTO_ENVIO_KEY, String(valor));
    }
  }

  public async getPedidos(): Promise<IPedido[]> {
    const items = await this.sharePointService.getListItems<SharePointPedidoItem>(
      this.pedidosListTitle,
      this.pedidoSelectFields
    );
    return items.map((item) => this.mapPedido(item));
  }

  public async addPedido(formData: IPedidoFormData): Promise<IPedido> {
    const createdItem = await this.sharePointService.addListItem<SharePointPedidoItem>(
      this.pedidosListTitle,
      this.mapPedidoPayload(formData),
      this.pedidoSelectFields
    );
    return this.mapPedido(createdItem);
  }

  public async updatePedido(id: number, formData: IPedidoFormData): Promise<void> {
    await this.sharePointService.updateListItem(this.pedidosListTitle, id, this.mapPedidoPayload(formData));
  }

  public async updateEstado(id: number, estado: string): Promise<void> {
    await this.sharePointService.updateListItem(this.pedidosListTitle, id, { EstadoPedido: estado });
  }

  public async deletePedido(id: number): Promise<void> {
    const existingItems = await this.getItemsByPedidoId(id);
    await Promise.all(existingItems.map((item) => this.sharePointService.deleteListItem(this.itemsListTitle, item.ID)));
    await this.sharePointService.deleteListItem(this.pedidosListTitle, id);
  }

  public async getItemsByPedidoId(pedidoId: number, formResponseId?: number): Promise<IPedidoItem[]> {
    const exprs: string[] = [];

    // Prioridad 1: relación PedidoFormResponseId ↔ FormResponseId de Pedidos
    if (formResponseId !== undefined && formResponseId !== null) {
      exprs.push(`PedidoFormResponseId eq ${formResponseId}`);
    }
    // Prioridad 2: lookup interno de SharePoint
    exprs.push(`PedidoId eq ${pedidoId}`);
    exprs.push(`Pedido eq ${pedidoId}`);
    exprs.push(`Pedido eq '${pedidoId}'`);

    for (const expr of exprs) {
      try {
        const items = await this.sharePointService.getListItemsFiltered<SharePointPedidoItemRow>(
          this.itemsListTitle,
          this.itemSelectFields,
          expr
        );
        return items.map((item) => this.mapPedidoItem(item));
      } catch (err) {
        console.warn(`[PedidoService] Filtro fallido (${expr}):`, err);
      }
    }

    return [];
  }

  public async saveItems(pedidoId: number, itemsForm: IPedidoItemFormData[], formResponseId?: number): Promise<void> {
    const existingItems = await this.getItemsByPedidoId(pedidoId, formResponseId);

    const existingIds = new Set(existingItems.map((i) => i.ID));
    const formIds = new Set(itemsForm.filter((i) => i.ID !== undefined).map((i) => i.ID as number));

    const toDelete = existingItems.filter((i) => !formIds.has(i.ID));
    await Promise.all(toDelete.map((i) => this.sharePointService.deleteListItem(this.itemsListTitle, i.ID)));

    const toUpdate = itemsForm.filter((i) => i.ID !== undefined && existingIds.has(i.ID as number));
    await Promise.all(
      toUpdate.map((i) =>
        this.sharePointService.updateListItem(this.itemsListTitle, i.ID as number, this.mapItemPayload(pedidoId, i, formResponseId))
      )
    );

    const toAdd = itemsForm.filter((i) => i.ID === undefined);
    await Promise.all(
      toAdd.map((i) =>
        this.sharePointService.addListItem(this.itemsListTitle, this.mapItemPayload(pedidoId, i, formResponseId))
      )
    );
  }

  private mapPedidoPayload(formData: IPedidoFormData): Record<string, unknown> {
    return {
      Title: formData.Title.trim(),
      NombreCompleto: formData.NombreCompleto.trim(),
      WhatsApp: formData.WhatsApp.trim(),
      MetodoEntrega: formData.MetodoEntrega,
      DireccionEntrega: formData.DireccionEntrega.trim(),
      FechaEntrega: formData.FechaEntrega || null,
      HorarioAproximado: formData.HorarioAproximado,
      MetodoPago: formData.MetodoPago,
      CubiertosDescartables: formData.CubiertosDescartables,
      Comentarios: formData.Comentarios.trim(),
      EstadoPedido: formData.EstadoPedido
    };
  }

  private mapItemPayload(pedidoId: number, item: IPedidoItemFormData, formResponseId?: number): Record<string, unknown> {
    const payload: Record<string, unknown> = {
      Title: item.Producto.trim(),
      Producto: item.Producto.trim(),
      Cantidad: item.Cantidad,
      Orden: item.Orden
    };

    if (formResponseId !== undefined && formResponseId !== null) {
      payload.FormResponseId = formResponseId;
    } else {
      payload.PedidoId = pedidoId;
    }

    return payload;
  }

  private mapPedido(item: SharePointPedidoItem): IPedido {
    return {
      ID: item.ID,
      Title: item.Title || '',
      FormResponseId: item.FormResponseId,
      SubmittedAt: item.SubmittedAt,
      NombreCompleto: item.NombreCompleto || '',
      WhatsApp: item.WhatsApp || '',
      MetodoEntrega: item.MetodoEntrega || '',
      DireccionEntrega: item.DireccionEntrega || '',
      FechaEntrega: item.FechaEntrega,
      HorarioAproximado: item.HorarioAproximado || '',
      MetodoPago: item.MetodoPago || '',
      CubiertosDescartables: item.CubiertosDescartables ?? false,
      Comentarios: item.Comentarios || '',
      EstadoPedido: item.EstadoPedido || 'Nuevo'
    };
  }

  private mapPedidoItem(item: SharePointPedidoItemRow): IPedidoItem {
    return {
      ID: item.ID,
      Title: item.Title || '',
      Pedido: String(item.Pedido ?? item.PedidoId ?? ''),
      Producto: item.Producto || '',
      Cantidad: item.Cantidad ?? 1,
      Orden: item.Orden ?? 0
    };
  }
}