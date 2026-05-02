import type { IPedido, IPedidoFormData } from '../../models/IPedido';
import type { IPedidoItem, IPedidoItemFormData } from '../../models/IPedidoItem';
import { SharePointService } from '../../services/SharePointService';

type SharePointPedidoItem = {
  ID: number;
  Title?: string;
  FormResponseId?: number;
  SubmittedAt?: string;
  CompletedAt?: string;
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
  Pedido?: string;
  Producto?: string;
  Cantidad?: number;
  Orden?: number;
};

export class PedidoService {
  private readonly sharePointService: SharePointService;

  private readonly pedidosListTitle: string = 'Pedidos';
  private readonly itemsListTitle: string = 'PedidoItems';

  private readonly pedidoSelectFields: string[] = [
    'ID', 'Title', 'FormResponseId', 'SubmittedAt', 'CompletedAt',
    'NombreCompleto', 'WhatsApp', 'MetodoEntrega', 'DireccionEntrega',
    'FechaEntrega', 'HorarioAproximado', 'MetodoPago', 'CubiertosDescartables',
    'Comentarios', 'EstadoPedido'
  ];

  private readonly itemSelectFields: string[] = [
    'ID', 'Title', 'Pedido', 'Producto', 'Cantidad', 'Orden'
  ];

  constructor(sharePointService: SharePointService) {
    this.sharePointService = sharePointService;
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

  public async deletePedido(id: number): Promise<void> {
    const existingItems = await this.getItemsByPedidoId(id);
    await Promise.all(existingItems.map((item) => this.sharePointService.deleteListItem(this.itemsListTitle, item.ID)));
    await this.sharePointService.deleteListItem(this.pedidosListTitle, id);
  }

  public async getItemsByPedidoId(pedidoId: number): Promise<IPedidoItem[]> {
    const items = await this.sharePointService.getListItemsFiltered<SharePointPedidoItemRow>(
      this.itemsListTitle,
      this.itemSelectFields,
      `Pedido eq '${pedidoId}'`
    );
    return items.map((item) => this.mapPedidoItem(item));
  }

  public async saveItems(pedidoId: number, itemsForm: IPedidoItemFormData[]): Promise<void> {
    const existingItems = await this.getItemsByPedidoId(pedidoId);

    const existingIds = new Set(existingItems.map((i) => i.ID));
    const formIds = new Set(itemsForm.filter((i) => i.ID !== undefined).map((i) => i.ID as number));

    // Delete items removed from form
    const toDelete = existingItems.filter((i) => !formIds.has(i.ID));
    await Promise.all(toDelete.map((i) => this.sharePointService.deleteListItem(this.itemsListTitle, i.ID)));

    // Update existing items
    const toUpdate = itemsForm.filter((i) => i.ID !== undefined && existingIds.has(i.ID as number));
    await Promise.all(
      toUpdate.map((i) =>
        this.sharePointService.updateListItem(this.itemsListTitle, i.ID as number, this.mapItemPayload(pedidoId, i))
      )
    );

    // Add new items
    const toAdd = itemsForm.filter((i) => i.ID === undefined);
    await Promise.all(
      toAdd.map((i) =>
        this.sharePointService.addListItem(this.itemsListTitle, this.mapItemPayload(pedidoId, i))
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

  private mapItemPayload(pedidoId: number, item: IPedidoItemFormData): Record<string, unknown> {
    return {
      Title: item.Producto.trim(),
      Pedido: String(pedidoId),
      Producto: item.Producto.trim(),
      Cantidad: item.Cantidad,
      Orden: item.Orden
    };
  }

  private mapPedido(item: SharePointPedidoItem): IPedido {
    return {
      ID: item.ID,
      Title: item.Title || '',
      FormResponseId: item.FormResponseId,
      SubmittedAt: item.SubmittedAt,
      CompletedAt: item.CompletedAt,
      NombreCompleto: item.NombreCompleto || '',
      WhatsApp: item.WhatsApp || '',
      MetodoEntrega: item.MetodoEntrega || '',
      DireccionEntrega: item.DireccionEntrega || '',
      FechaEntrega: item.FechaEntrega,
      HorarioAproximado: item.HorarioAproximado || '',
      MetodoPago: item.MetodoPago || '',
      CubiertosDescartables: item.CubiertosDescartables ?? false,
      Comentarios: item.Comentarios || '',
      EstadoPedido: item.EstadoPedido || 'Pendiente'
    };
  }

  private mapPedidoItem(item: SharePointPedidoItemRow): IPedidoItem {
    return {
      ID: item.ID,
      Title: item.Title || '',
      Pedido: item.Pedido || '',
      Producto: item.Producto || '',
      Cantidad: item.Cantidad ?? 1,
      Orden: item.Orden ?? 0
    };
  }
}
