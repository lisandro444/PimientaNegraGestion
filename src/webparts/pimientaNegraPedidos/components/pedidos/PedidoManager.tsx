import * as React from 'react';
import {
  DefaultButton,
  DatePicker,
  DayOfWeek,
  Dropdown,
  type IDropdownOption,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  Stack,
  Text,
  TextField,
  Toggle
} from '@fluentui/react';
import type { IPedido, IPedidoFormData } from '../../models/IPedido';
import type { IPedidoItem, IPedidoItemFormData } from '../../models/IPedidoItem';
import { SharePointService } from '../../services/SharePointService';
import { PedidoService } from './PedidoService';
import { PedidosListView } from './PedidosListView';
import { PedidoDetailView } from './PedidoDetailView';
import styles from './PedidoManager.module.scss';
import type { IPedidoManagerProps } from './IPedidoManagerProps';
import type { IPedidoManagerState, IPedidoFiltros } from './IPedidoManagerState';

// ── Constantes ──────────────────────────────────────────────────

const ESTADOS: IDropdownOption[] = [
  { key: 'Nuevo', text: 'Nuevo' },
  { key: 'Confirmado', text: 'Confirmado' },
  { key: 'En preparación', text: 'En preparación' },
  { key: 'Listo', text: 'Listo' },
  { key: 'Entregado', text: 'Entregado' },
  { key: 'Cancelado', text: 'Cancelado' }
];

const METODOS_ENTREGA: IDropdownOption[] = [
  { key: 'Retiro en local', text: 'Retiro en local' },
  { key: 'Delivery', text: 'Entrega a domicilio' }
];

const METODOS_PAGO: IDropdownOption[] = [
  { key: 'Efectivo', text: 'Efectivo' },
  { key: 'Transferencia', text: 'Transferencia' },
  { key: 'Mercado Pago', text: 'Mercado Pago' },
  { key: 'Tarjeta', text: 'Tarjeta' }
];

const HORARIOS: IDropdownOption[] = [
  { key: 'Mediodía', text: 'Mediodía (12-15hs)' },
  { key: 'Tarde', text: 'Tarde (19-21hs)' },
  { key: 'Noche', text: 'Noche (21-23hs)' }
];

const EMPTY_FORM: IPedidoFormData = {
  Title: '',
  NombreCompleto: '',
  WhatsApp: '',
  MetodoEntrega: 'Retiro en local',
  DireccionEntrega: '',
  FechaEntrega: '',
  HorarioAproximado: 'Mediodía',
  MetodoPago: 'Efectivo',
  CubiertosDescartables: false,
  Comentarios: '',
  EstadoPedido: 'Nuevo'
};

const EMPTY_FILTROS: IPedidoFiltros = {
  texto: '',
  estado: '',
  franja: '',
  metodo: ''
};

// ── Componente principal ────────────────────────────────────────

export default class PedidoManager extends React.Component<IPedidoManagerProps, IPedidoManagerState> {
  private readonly pedidoService: PedidoService;
  private itemsCache: Map<number, IPedidoItem[]> = new Map();
  private loadingSet: Set<number> = new Set();
  private expandedIds: number[] = [];

  constructor(props: IPedidoManagerProps) {
    super(props);
    this.pedidoService = new PedidoService(new SharePointService(props.context.pageContext));
    this.state = {
      view: 'lista',
      showDetailPanel: false,
      selectedPedido: undefined,
      pedidos: [],
      loading: true,
      error: '',
      success: '',
      filtros: { ...EMPTY_FILTROS },
      pedidosSeleccionados: [],
      soloHoy: false,
      pedidoItems: [],
      loadingItems: false,
      costoEnvio: this.pedidoService.getCostoEnvio(),
      savingEstado: false,
      isPanelOpen: false,
      saving: false,
      editingPedido: undefined,
      formData: { ...EMPTY_FORM },
      itemsForm: []
    };
  }

  public componentDidMount(): void {
    this.loadPedidos().catch((err) => {
      this.setState({
        loading: false,
        error: `Error al cargar pedidos: ${err instanceof Error ? err.message : String(err)}`
      });
    });
  }

  // ── Carga de datos ──────────────────────────────────────────────

  private readonly loadPedidos = async (): Promise<void> => {
    this.setState({ loading: true, error: '' });
    const pedidos = await this.pedidoService.getPedidos();
    const sorted = [...pedidos].sort((a, b) => b.ID - a.ID);
    this.setState({ pedidos: sorted, loading: false });
    // Pre-cargar los primeros 10 pedidos
    const toPreload = sorted.slice(0, 10);
    await Promise.all(toPreload.map((p) => this.ensureItemsLoaded(p.ID, p.FormResponseId)));
  };

  private async ensureItemsLoaded(pedidoId: number, formResponseId?: number): Promise<void> {
    if (this.itemsCache.has(pedidoId) || this.loadingSet.has(pedidoId)) return;
    this.loadingSet.add(pedidoId);
    this.forceUpdate();
    try {
      const items = await this.pedidoService.getItemsByPedidoId(pedidoId, formResponseId);
      this.itemsCache.set(pedidoId, items);
    } finally {
      this.loadingSet.delete(pedidoId);
      this.forceUpdate();
    }
  }

  private readonly getItemsForPedido = (pedidoId: number): IPedidoItem[] => {
    return this.itemsCache.get(pedidoId) ?? [];
  };

  // ── Render ──────────────────────────────────────────────────────

  public render(): React.ReactElement<IPedidoManagerProps> {
    const {
      selectedPedido, showDetailPanel, pedidos, loading, error, success,
      filtros, pedidosSeleccionados, soloHoy,
      pedidoItems, loadingItems, costoEnvio, savingEstado,
      isPanelOpen, saving, editingPedido, formData, itemsForm
    } = this.state;

    return (
      <div className={`${styles.managerRoot} ${this.props.hasTeamsContext ? styles.teamsHost : ''}`}>
        {error && (
          <div className={styles.globalMessageBar}>
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: '' })}>
              {error}
            </MessageBar>
          </div>
        )}
        {success && (
          <div className={styles.globalMessageBar}>
            <MessageBar messageBarType={MessageBarType.success} isMultiline={false} onDismiss={() => this.setState({ success: '' })}>
              {success}
            </MessageBar>
          </div>
        )}

        {/* Lista siempre visible */}
        <PedidosListView
            pedidos={pedidos}
            loading={loading}
            filtros={filtros}
            pedidosSeleccionados={pedidosSeleccionados}
            soloHoy={soloHoy}
            costoEnvio={costoEnvio}
            getItemsForPedido={this.getItemsForPedido}
            loadingItemsIds={Array.from(this.loadingSet)}
            expandedIds={this.expandedIds}
            onFiltroChange={this.handleFiltroChange}
            onToggleSeleccion={this.handleToggleSeleccion}
            onSeleccionarTodos={this.handleSeleccionarTodos}
            onLimpiarSeleccion={this.handleLimpiarSeleccion}
            onCambiarEstadoMasivo={(estado) => {
              this.handleCambiarEstadoMasivo(estado).catch(console.error);
            }}
            onVerDetalle={this.handleVerDetalle}
            onToggleExpand={this.handleToggleExpand}
            onSoloHoy={() => this.setState((s) => ({ soloHoy: !s.soloHoy }))}
            onCostoEnvioChange={(valor) => {
              this.pedidoService.setCostoEnvio(valor);
              this.setState({ costoEnvio: valor });
            }}
            onNuevoPedido={this.openNewPanel}
          />

        {/* Panel detalle (popup al hacer clic en card) */}
        {showDetailPanel && selectedPedido && (
          <Panel
            isOpen={showDetailPanel}
            onDismiss={() => this.setState({ showDetailPanel: false, selectedPedido: undefined })}
            type={PanelType.large}
            hasCloseButton
            closeButtonAriaLabel="Cerrar"
            isLightDismiss
            styles={{
              header: { display: 'none' },
              content: { padding: '0' },
              scrollableContent: { padding: '0' }
            }}
          >
            <PedidoDetailView
              pedido={selectedPedido}
              items={pedidoItems}
              loadingItems={loadingItems}
              costoEnvio={costoEnvio}
              savingEstado={savingEstado}
              onVolver={() => this.setState({ showDetailPanel: false, selectedPedido: undefined })}
              onCambiarEstado={(estado) => {
                this.handleCambiarEstado(selectedPedido.ID, estado).catch(console.error);
              }}
              onEditar={() => this.openEditPanel(selectedPedido)}
            />
          </Panel>
        )}

        {this.renderPanel(isPanelOpen, editingPedido, formData, itemsForm, saving)}
      </div>
    );
  }

  // ── Panel ABM ────────────────────────────────────────────────────

  private renderPanel(
    isPanelOpen: boolean,
    editingPedido: IPedido | undefined,
    formData: IPedidoFormData,
    itemsForm: IPedidoItemFormData[],
    saving: boolean
  ): React.ReactElement {
    return (
      <Panel
        isOpen={isPanelOpen}
        onDismiss={this.closePanel}
        type={PanelType.medium}
        headerText={editingPedido ? `Editar — ${editingPedido.Title}` : 'Nuevo pedido'}
        closeButtonAriaLabel="Cerrar"
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField label="N° / Título" required value={formData.Title}
            onChange={(_, v) => this.updateFormField('Title', v || '')} />
          <TextField label="Nombre completo" required value={formData.NombreCompleto}
            onChange={(_, v) => this.updateFormField('NombreCompleto', v || '')} />
          <TextField label="WhatsApp" value={formData.WhatsApp}
            onChange={(_, v) => this.updateFormField('WhatsApp', v || '')} />
          <Dropdown label="Estado" required selectedKey={formData.EstadoPedido} options={ESTADOS}
            onChange={(_, o) => this.updateFormField('EstadoPedido', o?.key as string || 'Nuevo')} />
          <Dropdown label="Método de entrega" required selectedKey={formData.MetodoEntrega} options={METODOS_ENTREGA}
            onChange={(_, o) => this.updateFormField('MetodoEntrega', o?.key as string || '')} />
          {formData.MetodoEntrega === 'Delivery' && (
            <TextField label="Dirección de entrega" value={formData.DireccionEntrega}
              onChange={(_, v) => this.updateFormField('DireccionEntrega', v || '')} />
          )}
          <DatePicker label="Fecha de entrega" firstDayOfWeek={DayOfWeek.Monday}
            placeholder="Seleccioná una fecha"
            value={formData.FechaEntrega ? new Date(formData.FechaEntrega) : undefined}
            onSelectDate={(d) => this.updateFormField('FechaEntrega', d ? d.toISOString() : '')}
            formatDate={(d) => d ? d.toLocaleDateString('es-AR') : ''} />
          <Dropdown label="Horario aproximado" selectedKey={formData.HorarioAproximado || null} options={HORARIOS}
            onChange={(_, o) => this.updateFormField('HorarioAproximado', o?.key as string || '')} />
          <Dropdown label="Método de pago" required selectedKey={formData.MetodoPago} options={METODOS_PAGO}
            onChange={(_, o) => this.updateFormField('MetodoPago', o?.key as string || '')} />
          <Toggle label="Cubiertos descartables" checked={formData.CubiertosDescartables}
            onText="Sí" offText="No"
            onChange={(_, checked) => this.updateFormField('CubiertosDescartables', !!checked)} />
          <TextField label="Comentarios" multiline rows={3} value={formData.Comentarios}
            onChange={(_, v) => this.updateFormField('Comentarios', v || '')} />

          <Text variant="mediumPlus" block className={styles.panelSectionTitle}>Productos</Text>
          {itemsForm.length === 0 && <Text className={styles.noProductos}>Sin productos cargados.</Text>}
          {itemsForm.map((item, idx) => (
            <div key={item.localId || String(item.ID)} className={styles.itemFormRow}>
              <TextField placeholder="Producto" value={item.Producto} styles={{ root: { flex: 2 } }}
                onChange={(_, v) => this.updateItemField(idx, 'Producto', v || '')} />
              <TextField placeholder="Cant." type="number" value={String(item.Cantidad)} styles={{ root: { width: 70 } }}
                onChange={(_, v) => this.updateItemField(idx, 'Cantidad', parseInt(v || '1', 10) || 1)} />

              <button className={styles.btnRemoveItem} onClick={() => this.removeItem(idx)} title="Quitar">✕</button>
            </div>
          ))}
          <button className={styles.btnAddItem} onClick={this.addItem}>+ Agregar producto</button>

          <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.panelFooter}>
            <PrimaryButton text={saving ? 'Guardando...' : 'Guardar'}
              onClick={() => { this.handleSave().catch(console.error); }} disabled={saving} />
            <DefaultButton text="Cancelar" onClick={this.closePanel} disabled={saving} />
          </Stack>
        </Stack>
      </Panel>
    );
  }

  // ── Handlers — navegación / lista ───────────────────────────────

  private readonly handleToggleExpand = (pedidoId: number): void => {
    if (this.expandedIds.includes(pedidoId)) {
      this.expandedIds = this.expandedIds.filter((id) => id !== pedidoId);
    } else {
      this.expandedIds = [...this.expandedIds, pedidoId];
      const pedido = this.state.pedidos.find((p) => p.ID === pedidoId);
      this.ensureItemsLoaded(pedidoId, pedido?.FormResponseId).catch(console.error);
    }
    this.forceUpdate();
  };

  private readonly handleFiltroChange = (partial: Partial<IPedidoFiltros>): void => {
    this.setState((s) => ({ filtros: { ...s.filtros, ...partial } }));
  };

  private readonly handleToggleSeleccion = (id: number): void => {
    this.setState((s) => ({
      pedidosSeleccionados: s.pedidosSeleccionados.includes(id)
        ? s.pedidosSeleccionados.filter((x) => x !== id)
        : [...s.pedidosSeleccionados, id]
    }));
  };

  private readonly handleSeleccionarTodos = (): void => {
    this.setState((s) => ({ pedidosSeleccionados: s.pedidos.map((p) => p.ID) }));
  };

  private readonly handleLimpiarSeleccion = (): void => {
    this.setState({ pedidosSeleccionados: [] });
  };

  private readonly handleCambiarEstadoMasivo = async (estado: string): Promise<void> => {
    const { pedidosSeleccionados } = this.state;
    if (pedidosSeleccionados.length === 0) return;
    try {
      this.setState({ savingEstado: true });
      await Promise.all(pedidosSeleccionados.map((id) => this.pedidoService.updateEstado(id, estado)));
      this.setState((s) => ({
        pedidos: s.pedidos.map((p) =>
          pedidosSeleccionados.includes(p.ID) ? { ...p, EstadoPedido: estado } : p
        ),
        pedidosSeleccionados: [],
        success: `${pedidosSeleccionados.length} pedido(s) actualizados a "${estado}".`
      }));
    } catch (err) {
      this.setState({ error: `Error al cambiar estado: ${err instanceof Error ? err.message : String(err)}` });
    } finally {
      this.setState({ savingEstado: false });
    }
  };

  private readonly handleVerDetalle = (pedido: IPedido): void => {
    const cached = this.itemsCache.get(pedido.ID);
    this.setState({
      showDetailPanel: true,
      selectedPedido: pedido,
      pedidoItems: cached ?? [],
      loadingItems: !cached
    });
    if (!cached) {
      this.pedidoService.getItemsByPedidoId(pedido.ID, pedido.FormResponseId)
        .then((items) => {
          this.itemsCache.set(pedido.ID, items);
          this.setState({ pedidoItems: items, loadingItems: false });
        })
        .catch((err) => this.setState({
          loadingItems: false,
          error: `Error al cargar productos: ${err instanceof Error ? err.message : String(err)}`
        }));
    }
  };

  private readonly handleCambiarEstado = async (pedidoId: number, estado: string): Promise<void> => {
    this.setState({ savingEstado: true });
    try {
      await this.pedidoService.updateEstado(pedidoId, estado);
      this.setState((s) => ({
        pedidos: s.pedidos.map((p) => p.ID === pedidoId ? { ...p, EstadoPedido: estado } : p),
        selectedPedido: s.selectedPedido?.ID === pedidoId
          ? { ...s.selectedPedido, EstadoPedido: estado }
          : s.selectedPedido,
        success: `Estado actualizado a "${estado}".`
      }));
    } catch (err) {
      this.setState({ error: `Error al cambiar estado: ${err instanceof Error ? err.message : String(err)}` });
    } finally {
      this.setState({ savingEstado: false });
    }
  };

  // ── Handlers — ABM ──────────────────────────────────────────────

  private readonly openNewPanel = (): void => {
    this.setState({
      isPanelOpen: true, editingPedido: undefined, error: '', success: '',
      formData: { ...EMPTY_FORM }, itemsForm: []
    });
  };

  private openEditPanel(pedido: IPedido): void {
    const cachedItems = this.itemsCache.get(pedido.ID) ?? [];
    this.setState({
      isPanelOpen: true,
      editingPedido: pedido,
      error: '', success: '',
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
      itemsForm: cachedItems.map((i) => ({
        localId: String(i.ID), ID: i.ID,
        Producto: i.Producto, Cantidad: i.Cantidad,
        Orden: i.Orden
      }))
    });
  }

  private readonly closePanel = (): void => {
    if (!this.state.saving) this.setState({ isPanelOpen: false });
  };

  private updateFormField(field: keyof IPedidoFormData, value: string | boolean): void {
    this.setState((s) => ({ formData: { ...s.formData, [field]: value } }));
  }

  private updateItemField(index: number, field: keyof IPedidoItemFormData, value: string | number): void {
    this.setState((s) => {
      const updated = [...s.itemsForm];
      updated[index] = { ...updated[index], [field]: value };
      return { itemsForm: updated };
    });
  }

  private readonly addItem = (): void => {
    this.setState((s) => ({
      itemsForm: [...s.itemsForm, {
        localId: `new-${Date.now()}`,
        Producto: '', Cantidad: 1,
        Orden: s.itemsForm.length + 1
      }]
    }));
  };

  private removeItem(index: number): void {
    this.setState((s) => {
      const updated = [...s.itemsForm];
      updated.splice(index, 1);
      return { itemsForm: updated };
    });
  }

  private validateForm(): string | undefined {
    const { formData, itemsForm } = this.state;
    if (!formData.Title.trim()) return 'El número/título del pedido es requerido.';
    if (!formData.NombreCompleto.trim()) return 'El nombre del cliente es requerido.';
    for (const item of itemsForm) {
      if (!item.Producto.trim()) return 'Todos los productos deben tener nombre.';
      if (item.Cantidad < 1) return 'La cantidad debe ser al menos 1.';
    }
    return undefined;
  }

  private async handleSave(): Promise<void> {
    const { editingPedido, formData, itemsForm } = this.state;
    const err = this.validateForm();
    if (err) { this.setState({ error: err }); return; }

    try {
      this.setState({ saving: true, error: '', success: '' });
      let pedidoId: number;
      let formResponseId: number | undefined;

      if (editingPedido) {
        await this.pedidoService.updatePedido(editingPedido.ID, formData);
        pedidoId = editingPedido.ID;
        formResponseId = editingPedido.FormResponseId;
      } else {
        const created = await this.pedidoService.addPedido(formData);
        pedidoId = created.ID;
        formResponseId = created.FormResponseId;
      }

      await this.pedidoService.saveItems(pedidoId, itemsForm, formResponseId);
      this.itemsCache.delete(pedidoId);

      this.setState({
        isPanelOpen: false, editingPedido: undefined,
        formData: { ...EMPTY_FORM }, itemsForm: [],
        success: editingPedido ? 'Pedido actualizado.' : 'Pedido creado.'
      });

      await this.loadPedidos();

      if (this.state.view === 'detalle' && this.state.selectedPedido?.ID === pedidoId) {
        const updatedItems = await this.pedidoService.getItemsByPedidoId(pedidoId, formResponseId);
        this.itemsCache.set(pedidoId, updatedItems);
        const updatedPedido = this.state.pedidos.find((p) => p.ID === pedidoId);
        this.setState({ pedidoItems: updatedItems, selectedPedido: updatedPedido });
      }
    } catch (saveErr) {
      this.setState({ error: `Error al guardar: ${saveErr instanceof Error ? saveErr.message : String(saveErr)}` });
    } finally {
      this.setState({ saving: false });
    }
  }
}