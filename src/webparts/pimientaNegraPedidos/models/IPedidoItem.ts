export interface IPedidoItem {
  ID: number;
  Title: string;
  Pedido: string;
  Producto: string;
  Cantidad: number;
  Orden: number;
}

export interface IPedidoItemFormData {
  localId?: string;
  ID?: number;
  Producto: string;
  Cantidad: number;
  Orden: number;
}
