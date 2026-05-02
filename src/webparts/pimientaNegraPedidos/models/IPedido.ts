export interface IPedido {
  ID: number;
  Title: string;
  FormResponseId?: number;
  SubmittedAt?: string;
  CompletedAt?: string;
  NombreCompleto: string;
  WhatsApp: string;
  MetodoEntrega: string;
  DireccionEntrega: string;
  FechaEntrega?: string;
  HorarioAproximado: string;
  MetodoPago: string;
  CubiertosDescartables: boolean;
  Comentarios: string;
  EstadoPedido: string;
}

export interface IPedidoFormData {
  Title: string;
  NombreCompleto: string;
  WhatsApp: string;
  MetodoEntrega: string;
  DireccionEntrega: string;
  FechaEntrega: string;
  HorarioAproximado: string;
  MetodoPago: string;
  CubiertosDescartables: boolean;
  Comentarios: string;
  EstadoPedido: string;
}
