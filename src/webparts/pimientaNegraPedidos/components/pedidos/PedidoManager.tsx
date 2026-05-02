import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  type IColumn,
  IconButton,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  SelectionMode,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  type IDropdownOption,
  DatePicker,
  DayOfWeek
} from '@fluentui/react';
import type { IPedido } from '../../models/IPedido';
import type { IPedidoItemFormData } from '../../models/IPedidoItem';
import { SharePointService } from '../../services/SharePointService';
import { PedidoService } from './PedidoService';
import styles from './PedidoManager.module.scss';
import type { IPedidoManagerProps } from './IPedidoManagerProps';
import type { IPedidoManagerState } from './IPedidoManagerState';

const ESTADOS: IDropdownOption[] = [
  { key: 'Pendiente', text: 'Pendiente' },
  { key: 'En preparación', text: 'En preparación' },
  { key: 'Listo', text: 'Listo' },
  { key: 'Entregado', text: 'Entregado' },
  { key: 'Cancelado', text: 'Cancelado' }
];

const METODOS_ENTREGA: IDropdownOption[] = [
  { key: 'Retiro en local', text: 'Retiro en local' },
  { key: 'Delivery', text: 'Delivery' }
];

const METODOS_PAGO: IDropdownOption[] = [
  { key: 'Efectivo', text: 'Efectivo' },
  { key: 'Transferencia', text: 'Transferencia' },
  { key: 'Mercado Pago', text: 'Mercado Pago' },
  { key: 'Tarjeta', text: 'Tarjeta' }
];

const HORARIOS: IDropdownOption[] = [
  { key: '12:00 - 13:00', text: '12:00 - 13:00' },
  { key: '13:00 - 14:00', text: '13:00 - 14:00' },
  { key: '14:00 - 15:00', text: '14:00 - 15:00' },
  { key: '19:00 - 20:00', text: '19:00 - 20:00' },
  { key: '20:00 - 21:00', text: '20:00 - 21:00' },
  { key: '21:00 - 22:00', text: '21:00 - 22:00' }
];

const EMPTY_FORM = {
  Title: '',
  NombreCompleto: '',
  WhatsApp: '',
  MetodoEntrega: 'Retiro en local',
  DireccionEntrega: '',
  FechaEntrega: '',
  HorarioAproximado: '',
  MetodoPago: 'Efectivo',
  CubiertosDescartables: false,
  Comentarios: '',
  EstadoPedido: 'Pendiente'
};

function getEstadoBadgeClass(estado: string): string {
  switch (estado) {
    case 'En preparación': return styles.estadoEnPreparacion;
    case 'Listo': return styles.estadoListo;
    case 'Entregado': return styles.estadoEntregado;
    case 'Cancelado': return styles.estadoCancelado;
    default: return styles.estadoPendiente;
  }
}

function formatDate(isoDate?: string): string {
  if (!isoDate) return '-';
  const d = new Date(isoDate);
  return d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export default class PedidoManager extends React.Component<IPedidoManagerProps, IPedidoManagerState> {
  private readonly pedidoService: PedidoService;

  constructor(props: IPedidoManagerProps) {
    super(props);

    this.pedidoService = new PedidoService(new SharePointService(props.context.pageContext));
    this.state = {
      pedidos: [],
      loading: true,
      saving: false,
      error: '',
      success: '',
      isPanelOpen: false,
      editingPedido: undefined,
      formData: { ...EMPTY_FORM },
      itemsForm: [],
      loadingItems: false
    };
  }

  public componentDidMount(): void {
    this.loadPedidos().catch((err) => {
      this.setState({
        loading: false,
        error: `Error al cargar pedidos: ${err instanceof Error ? err.message : 'Error desconocido'}`
      });
    });
  }

  public render(): React.ReactElement<IPedidoManagerProps> {
    const { pedidos, loading, error, success, isPanelOpen, editingPedido, formData, saving, itemsForm, loadingItems } = this.state;

    const columns: IColumn[] = [
      {
        key: 'title',
        name: 'N° Pedido',
        fieldName: 'Title',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: IPedido) => <span className={styles.clienteName}>{item.Title}</span>
      },
      {
        key: 'cliente',
        name: 'Cliente',
        fieldName: 'NombreCompleto',
        minWidth: 160,
        isResizable: true
      },
      {
        key: 'whatsapp',
        name: 'WhatsApp',
        fieldName: 'WhatsApp',
        minWidth: 120,
        isResizable: true
      },
      {
        key: 'entrega',
        name: 'Entrega',
        fieldName: 'MetodoEntrega',
        minWidth: 120,
        isResizable: true
      },
      {
        key: 'fecha',
        name: 'Fecha entrega',
        minWidth: 110,
        isResizable: true,
        onRender: (item: IPedido) => <span>{formatDate(item.FechaEntrega)}</span>
      },
      {
        key: 'pago',
        name: 'Pago',
        fieldName: 'MetodoPago',
        minWidth: 110,
        isResizable: true
      },
      {
        key: 'estado',
        name: 'Estado',
        minWidth: 130,
        isResizable: true,
        onRender: (item: IPedido) => (
          <span className={`${styles.statusBadge} ${getEstadoBadgeClass(item.EstadoPedido)}`}>
            {item.EstadoPedido}
          </span>
        )
      },
      {
        key: 'actions',
        name: 'Acciones',
        minWidth: 90,
        maxWidth: 100,
        onRender: (item: IPedido) => (
          <div className={styles.actionButtons}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Editar"
              ariaLabel="Editar"
              onClick={() => {
                this.openEditPanel(item).catch((err) => console.error('Error abriendo panel:', err));
              }}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Eliminar"
              ariaLabel="Eliminar"
              onClick={() => {
                this.handleDelete(item).catch((err) => console.error('Error eliminando:', err));
              }}
            />
          </div>
        )
      }
    ];

    return (
      <div className={styles.pedidosShell}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="end" className={styles.pageHeader}>
          <div>
            <Text variant="xxLarge" block className={styles.pageTitle}>
              Pedidos
            </Text>
            <Text variant="medium" className={styles.pageSubtitle}>
              Gestión de pedidos — Pimienta Negra Cocina
            </Text>
          </div>
          <PrimaryButton
            text="Nuevo pedido"
            iconProps={{ iconName: 'Add' }}
            onClick={this.openNewPanel}
            className={styles.addButton}
          />
        </Stack>

        {error && (
          <div className={styles.messageBar}>
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: '' })}>
              {error}
            </MessageBar>
          </div>
        )}

        {success && (
          <div className={styles.messageBar}>
            <MessageBar messageBarType={MessageBarType.success} isMultiline={false} onDismiss={() => this.setState({ success: '' })}>
              {success}
            </MessageBar>
          </div>
        )}

        <div className={styles.tableCard}>
          {loading ? (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label="Cargando pedidos..." />
            </div>
          ) : pedidos.length === 0 ? (
            <div className={styles.emptyState}>
              <Text>No hay pedidos registrados. Creá el primero.</Text>
            </div>
          ) : (
            <DetailsList
              items={pedidos}
              columns={columns}
              selectionMode={SelectionMode.none}
              className={styles.pedidosList}
            />
          )}
        </div>

        <Panel
          isOpen={isPanelOpen}
          onDismiss={this.closePanel}
          type={PanelType.medium}
          headerText={editingPedido ? `Editar pedido — ${editingPedido.Title}` : 'Nuevo pedido'}
          closeButtonAriaLabel="Cerrar"
        >
          <Stack tokens={{ childrenGap: 12 }}>
            {/* ── Datos del pedido ─────────────────────────── */}
            <TextField
              label="N° / Título del pedido"
              required
              value={formData.Title}
              onChange={(_, v) => this.updateFormField('Title', v || '')}
            />
            <TextField
              label="Nombre completo"
              required
              value={formData.NombreCompleto}
              onChange={(_, v) => this.updateFormField('NombreCompleto', v || '')}
            />
            <TextField
              label="WhatsApp"
              value={formData.WhatsApp}
              onChange={(_, v) => this.updateFormField('WhatsApp', v || '')}
            />
            <Dropdown
              label="Estado"
              required
              selectedKey={formData.EstadoPedido}
              options={ESTADOS}
              onChange={(_, option) => this.updateFormField('EstadoPedido', option?.key as string || 'Pendiente')}
            />
            <Dropdown
              label="Método de entrega"
              required
              selectedKey={formData.MetodoEntrega}
              options={METODOS_ENTREGA}
              onChange={(_, option) => this.updateFormField('MetodoEntrega', option?.key as string || '')}
            />
            {formData.MetodoEntrega === 'Delivery' && (
              <TextField
                label="Dirección de entrega"
                value={formData.DireccionEntrega}
                onChange={(_, v) => this.updateFormField('DireccionEntrega', v || '')}
              />
            )}
            <DatePicker
              label="Fecha de entrega"
              firstDayOfWeek={DayOfWeek.Monday}
              placeholder="Seleccioná una fecha"
              value={formData.FechaEntrega ? new Date(formData.FechaEntrega) : undefined}
              onSelectDate={(date) => this.updateFormField('FechaEntrega', date ? date.toISOString() : '')}
              formatDate={(date) => (date ? date.toLocaleDateString('es-AR') : '')}
            />
            <Dropdown
              label="Horario aproximado"
              selectedKey={formData.HorarioAproximado || null}
              options={HORARIOS}
              onChange={(_, option) => this.updateFormField('HorarioAproximado', option?.key as string || '')}
            />
            <Dropdown
              label="Método de pago"
              required
              selectedKey={formData.MetodoPago}
              options={METODOS_PAGO}
              onChange={(_, option) => this.updateFormField('MetodoPago', option?.key as string || '')}
            />
            <Toggle
              label="Cubiertos descartables"
              checked={formData.CubiertosDescartables}
              onText="Sí"
              offText="No"
              onChange={(_, checked) => this.updateFormField('CubiertosDescartables', !!checked)}
            />
            <TextField
              label="Comentarios"
              multiline
              rows={3}
              value={formData.Comentarios}
              onChange={(_, v) => this.updateFormField('Comentarios', v || '')}
            />

            {/* ── Ítems del pedido ─────────────────────────── */}
            <Text variant="mediumPlus" block className={styles.sectionTitle}>
              Productos del pedido
            </Text>

            {loadingItems ? (
              <Spinner size={SpinnerSize.small} label="Cargando productos..." />
            ) : (
              <>
                {itemsForm.length === 0 && (
                  <Text className={styles.noItems}>Sin productos cargados.</Text>
                )}
                {itemsForm.map((item, index) => (
                  <div key={item.localId || String(item.ID)} className={styles.itemRow}>
                    <div className={styles.itemFieldProducto}>
                      <TextField
                        label={index === 0 ? 'Producto' : undefined}
                        placeholder="Nombre del producto"
                        value={item.Producto}
                        onChange={(_, v) => this.updateItemField(index, 'Producto', v || '')}
                      />
                    </div>
                    <div className={styles.itemFieldCantidad}>
                      <TextField
                        label={index === 0 ? 'Cant.' : undefined}
                        type="number"
                        value={String(item.Cantidad)}
                        onChange={(_, v) => this.updateItemField(index, 'Cantidad', parseInt(v || '1', 10) || 1)}
                      />
                    </div>
                    <div className={styles.itemFieldOrden}>
                      <TextField
                        label={index === 0 ? 'Orden' : undefined}
                        type="number"
                        value={String(item.Orden)}
                        onChange={(_, v) => this.updateItemField(index, 'Orden', parseInt(v || '0', 10) || 0)}
                      />
                    </div>
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Quitar producto"
                      ariaLabel="Quitar producto"
                      onClick={() => this.removeItem(index)}
                    />
                  </div>
                ))}
                <DefaultButton
                  text="Agregar producto"
                  iconProps={{ iconName: 'Add' }}
                  onClick={this.addItem}
                  className={styles.addItemButton}
                />
              </>
            )}

            {/* ── Botones del panel ────────────────────────── */}
            <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.panelFooter}>
              <PrimaryButton
                text={saving ? 'Guardando...' : 'Guardar'}
                onClick={() => {
                  this.handleSave().catch((err) => console.error('Error guardando:', err));
                }}
                disabled={saving}
              />
              <DefaultButton text="Cancelar" onClick={this.closePanel} disabled={saving} />
            </Stack>
          </Stack>
        </Panel>
      </div>
    );
  }

  // ── Carga ───────────────────────────────────────────────────────

  private readonly loadPedidos = async (): Promise<void> => {
    this.setState({ loading: true, error: '' });
    const items = await this.pedidoService.getPedidos();
    const sorted = [...items].sort((a, b) => b.ID - a.ID);
    this.setState({ pedidos: sorted, loading: false });
  };

  // ── Panel ───────────────────────────────────────────────────────

  private readonly openNewPanel = (): void => {
    this.setState({
      editingPedido: undefined,
      isPanelOpen: true,
      error: '',
      success: '',
      formData: { ...EMPTY_FORM },
      itemsForm: []
    });
  };

  private async openEditPanel(pedido: IPedido): Promise<void> {
    this.setState({
      editingPedido: pedido,
      isPanelOpen: true,
      error: '',
      success: '',
      loadingItems: true,
      formData: {
        Title: pedido.Title,
        NombreCompleto: pedido.NombreCompleto,
        WhatsApp: pedido.WhatsApp,
        MetodoEntrega: pedido.MetodoEntrega,
        DireccionEntrega: pedido.DireccionEntrega,
        FechaEntrega: pedido.FechaEntrega || '',
        HorarioAproximado: pedido.HorarioAproximado,
        MetodoPago: pedido.MetodoPago,
        CubiertosDescartables: pedido.CubiertosDescartables,
        Comentarios: pedido.Comentarios,
        EstadoPedido: pedido.EstadoPedido
      },
      itemsForm: []
    });

    const items = await this.pedidoService.getItemsByPedidoId(pedido.ID);
    const itemsForm: IPedidoItemFormData[] = items
      .sort((a, b) => a.Orden - b.Orden)
      .map((i) => ({
        localId: String(i.ID),
        ID: i.ID,
        Producto: i.Producto,
        Cantidad: i.Cantidad,
        Orden: i.Orden
      }));

    this.setState({ itemsForm, loadingItems: false });
  }

  private readonly closePanel = (): void => {
    if (!this.state.saving) {
      this.setState({ isPanelOpen: false });
    }
  };

  // ── Formulario ─────────────────────────────────────────────────

  private updateFormField(
    field: keyof typeof EMPTY_FORM,
    value: string | boolean
  ): void {
    this.setState((s) => ({ formData: { ...s.formData, [field]: value } }));
  }

  private updateItemField(
    index: number,
    field: keyof IPedidoItemFormData,
    value: string | number
  ): void {
    this.setState((s) => {
      const updated = [...s.itemsForm];
      updated[index] = { ...updated[index], [field]: value };
      return { itemsForm: updated };
    });
  }

  private readonly addItem = (): void => {
    this.setState((s) => ({
      itemsForm: [
        ...s.itemsForm,
        {
          localId: `new-${Date.now()}`,
          Producto: '',
          Cantidad: 1,
          Orden: s.itemsForm.length + 1
        }
      ]
    }));
  };

  private removeItem(index: number): void {
    this.setState((s) => {
      const updated = [...s.itemsForm];
      updated.splice(index, 1);
      return { itemsForm: updated };
    });
  }

  // ── Validación ──────────────────────────────────────────────────

  private validateForm(): string | undefined {
    const { formData, itemsForm } = this.state;

    if (!formData.Title.trim()) return 'El número/título del pedido es requerido.';
    if (!formData.NombreCompleto.trim()) return 'El nombre del cliente es requerido.';
    if (!formData.MetodoEntrega) return 'Seleccioná un método de entrega.';
    if (!formData.MetodoPago) return 'Seleccioná un método de pago.';
    if (!formData.EstadoPedido) return 'Seleccioná un estado para el pedido.';

    for (const item of itemsForm) {
      if (!item.Producto.trim()) return 'Todos los productos deben tener nombre.';
      if (item.Cantidad < 1) return 'La cantidad debe ser al menos 1.';
    }

    return undefined;
  }

  // ── CRUD ────────────────────────────────────────────────────────

  private async handleSave(): Promise<void> {
    const { editingPedido, formData, itemsForm } = this.state;
    const validationError = this.validateForm();

    if (validationError) {
      this.setState({ error: validationError });
      return;
    }

    try {
      this.setState({ saving: true, error: '', success: '' });

      let pedidoId: number;

      if (editingPedido) {
        await this.pedidoService.updatePedido(editingPedido.ID, formData);
        pedidoId = editingPedido.ID;
        this.setState({ success: 'Pedido actualizado correctamente.' });
      } else {
        const created = await this.pedidoService.addPedido(formData);
        pedidoId = created.ID;
        this.setState({ success: 'Pedido creado correctamente.' });
      }

      await this.pedidoService.saveItems(pedidoId, itemsForm);

      this.setState({
        isPanelOpen: false,
        editingPedido: undefined,
        formData: { ...EMPTY_FORM },
        itemsForm: []
      });

      await this.loadPedidos();
    } catch (err) {
      this.setState({
        error: `Error al guardar el pedido: ${err instanceof Error ? err.message : 'Error desconocido'}`
      });
    } finally {
      this.setState({ saving: false });
    }
  }

  private async handleDelete(pedido: IPedido): Promise<void> {
    const confirmed = window.confirm(
      `¿Eliminás el pedido "${pedido.Title}" de ${pedido.NombreCompleto}? Esta acción también borrará todos sus productos.`
    );
    if (!confirmed) return;

    try {
      this.setState({ error: '', success: '' });
      await this.pedidoService.deletePedido(pedido.ID);
      this.setState({ success: 'Pedido eliminado correctamente.' });
      await this.loadPedidos();
    } catch (err) {
      this.setState({
        error: `Error al eliminar el pedido: ${err instanceof Error ? err.message : 'Error desconocido'}`
      });
    }
  }
}
