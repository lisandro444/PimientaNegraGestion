import * as React from 'react';
import {
  Checkbox,
  Dropdown,
  type IDropdownOption,
  Spinner,
  SpinnerSize,
  Text,
  TextField
} from '@fluentui/react';
import type { IPedido } from '../../models/IPedido';
import type { IPedidoItem } from '../../models/IPedidoItem';
import type { IPedidoFiltros } from './IPedidoManagerState';
import styles from './PedidoManager.module.scss';

export interface IPedidosListViewProps {
  pedidos: IPedido[];
  loading: boolean;
  filtros: IPedidoFiltros;
  pedidosSeleccionados: number[];
  soloHoy: boolean;
  costoEnvio: number;
  getItemsForPedido: (pedidoId: number) => IPedidoItem[];
  loadingItemsIds: number[];
  onFiltroChange: (filtros: Partial<IPedidoFiltros>) => void;
  onToggleSeleccion: (id: number) => void;
  onSeleccionarTodos: () => void;
  onLimpiarSeleccion: () => void;
  onCambiarEstadoMasivo: (estado: string) => void;
  onVerDetalle: (pedido: IPedido) => void;
  onToggleExpand: (pedidoId: number) => void;
  expandedIds: number[];
  onSoloHoy: () => void;
  onCostoEnvioChange: (valor: number) => void;
  onNuevoPedido: () => void;
}

const ESTADOS_OPCIONES: IDropdownOption[] = [
  { key: '', text: 'Todos los estados' },
  { key: 'Nuevo', text: 'Nuevo' },
  { key: 'Confirmado', text: 'Confirmado' },
  { key: 'Entregado', text: 'Entregado' },
  { key: 'Cancelado', text: 'Cancelado' }
];

const FRANJAS_OPCIONES: IDropdownOption[] = [
  { key: '', text: 'Toda franja' },
  { key: 'Mediodía', text: 'Mediodía' },
  { key: 'Tarde', text: 'Tarde' },
  { key: 'Noche', text: 'Noche' }
];

const ESTADOS_BADGE: Record<string, string> = {
  'Nuevo': styles.estadoNuevo,
  'Confirmado': styles.estadoConfirmado,
  'Entregado': styles.estadoEntregado,
  'Cancelado': styles.estadoCancelado
};

const ESTADO_FLOW = ['Nuevo', 'Confirmado', 'Entregado'];

function getFranja(horario: string): string {
  if (!horario) return '';
  const h = horario.toLowerCase();
  if (h.includes('12') || h.includes('13') || h.includes('14') || h.includes('mediod')) return 'Mediodía';
  if (h.includes('19') || h.includes('20') || h.includes('21') || h.includes('tarde')) return 'Tarde';
  if (h.includes('22') || h.includes('23') || h.includes('noche')) return 'Noche';
  return horario;
}

function formatFecha(isoDate?: string): string {
  if (!isoDate) return '-';
  const d = new Date(isoDate);
  return d.toLocaleDateString('es-AR', { day: '2-digit', month: 'short', year: 'numeric' });
}

function formatMoney(n: number): string {
  return `$ ${n.toLocaleString('es-AR')}`;
}

export const PedidosListView: React.FC<IPedidosListViewProps> = (props) => {
  const {
    pedidos, loading, filtros, pedidosSeleccionados, soloHoy, costoEnvio,
    getItemsForPedido, loadingItemsIds, onFiltroChange, onToggleSeleccion,
    onSeleccionarTodos, onLimpiarSeleccion, onCambiarEstadoMasivo,
    onVerDetalle, onToggleExpand, expandedIds, onSoloHoy, onCostoEnvioChange, onNuevoPedido
  } = props;

  const [envioEditing, setEnvioEditing] = React.useState(false);
  const [envioInput, setEnvioInput] = React.useState('');

  const handleEnvioClick = (): void => {
    setEnvioInput(String(costoEnvio));
    setEnvioEditing(true);
  };

  const handleEnvioBlur = (): void => {
    const val = parseInt(envioInput, 10);
    if (!isNaN(val) && val >= 0) onCostoEnvioChange(val);
    setEnvioEditing(false);
  };

  const handleEnvioKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter') handleEnvioBlur();
    if (e.key === 'Escape') setEnvioEditing(false);
  };

  // ── Filtrado ──────────────────────────────────────────────────
  const pedidosFiltrados = React.useMemo(() => {
    let result = [...pedidos];

    if (soloHoy) {
      const hoy = new Date().toDateString();
      result = result.filter((p) => p.FechaEntrega && new Date(p.FechaEntrega).toDateString() === hoy);
    }

    if (filtros.texto) {
      const q = filtros.texto.toLowerCase();
      result = result.filter(
        (p) =>
          p.NombreCompleto.toLowerCase().includes(q) ||
          p.WhatsApp.toLowerCase().includes(q) ||
          p.Title.toLowerCase().includes(q) ||
          getItemsForPedido(p.ID).some((i) => i.Producto.toLowerCase().includes(q))
      );
    }

    if (filtros.estado) {
      result = result.filter((p) => p.EstadoPedido === filtros.estado);
    }

    if (filtros.franja) {
      result = result.filter((p) => getFranja(p.HorarioAproximado) === filtros.franja);
    }

    if (filtros.metodo === 'Delivery') {
      result = result.filter((p) => p.MetodoEntrega === 'Delivery' || p.MetodoEntrega === 'Entrega a domicilio');
    } else if (filtros.metodo === 'Retiro') {
      result = result.filter((p) => p.MetodoEntrega === 'Retiro en local');
    }

    return result;
  }, [pedidos, filtros, soloHoy]);

  const todosSeleccionados =
    pedidosFiltrados.length > 0 && pedidosFiltrados.every((p) => pedidosSeleccionados.includes(p.ID));

  return (
    <div className={styles.listViewShell}>

      {/* ── Header ─────────────────────────────────────────── */}
      <div className={styles.listHeader}>
        <div>
          <Text variant="xxLarge" block className={styles.pageTitle}>Pedidos</Text>
          <Text variant="small" className={styles.pageSubtitle}>
            {pedidosFiltrados.length} de {pedidos.length} pedidos
          </Text>
        </div>
        <div className={styles.headerActions}>
          {/* Envío inline editable */}
          <div className={styles.envioInline}>
            <span className={styles.envioInlineIcon}>🚚</span>
            <span className={styles.envioInlineLabel}>Envío:</span>
            {envioEditing ? (
              <input
                className={styles.envioInlineInput}
                type="number"
                min={0}
                value={envioInput}
                autoFocus
                onChange={(e) => setEnvioInput(e.target.value)}
                onBlur={handleEnvioBlur}
                onKeyDown={handleEnvioKeyDown}
              />
            ) : (
              <button className={styles.envioInlineValue} onClick={handleEnvioClick} title="Clic para editar">
                $ {costoEnvio.toLocaleString('es-AR')}
              </button>
            )}
          </div>
          <button
            className={`${styles.btnSecondary} ${soloHoy ? styles.btnActive : ''}`}
            onClick={onSoloHoy}
          >
            📅 Solo hoy
          </button>
          <button
            className={`${styles.btnSecondary} ${pedidosSeleccionados.length > 0 ? styles.btnActive : ''}`}
            onClick={pedidosSeleccionados.length > 0 ? onLimpiarSeleccion : onSeleccionarTodos}
          >
            ☑ Seleccionar varios
          </button>
        </div>
      </div>

      {/* ── Barra de filtros ────────────────────────────────── */}
      <div className={styles.filtrosCard}>
        <div className={styles.filtrosRow}>
          <div className={styles.filtroSearch}>
            <TextField
              placeholder="Buscar por cliente, WhatsApp, N° o producto"
              iconProps={{ iconName: 'Search' }}
              value={filtros.texto}
              onChange={(_, v) => onFiltroChange({ texto: v || '' })}
              borderless={false}
            />
          </div>
          <Dropdown
            placeholder="Todos los estados"
            selectedKey={filtros.estado || null}
            options={ESTADOS_OPCIONES}
            onChange={(_, o) => onFiltroChange({ estado: (o?.key as string) || '' })}
            styles={{ root: { width: 170 } }}
          />
          <Dropdown
            placeholder="Toda franja"
            selectedKey={filtros.franja || null}
            options={FRANJAS_OPCIONES}
            onChange={(_, o) => onFiltroChange({ franja: (o?.key as string) || '' })}
            styles={{ root: { width: 140 } }}
          />
        </div>
        <div className={styles.filtrosMetodoRow}>
          <span className={styles.filtroLabel}>Método:</span>
          {(['Todos', 'Entrega a domicilio', 'Retiro en local'] as const).map((m) => {
            const key = m === 'Todos' ? '' : m === 'Entrega a domicilio' ? 'Delivery' : 'Retiro';
            return (
              <button
                key={m}
                className={`${styles.btnMetodo} ${filtros.metodo === key ? styles.btnMetodoActive : ''}`}
                onClick={() => onFiltroChange({ metodo: key })}
              >
                {m}
              </button>
            );
          })}
          <button
            className={styles.btnLimpiar}
            onClick={() => onFiltroChange({ texto: '', estado: '', franja: '', metodo: '' })}
          >
            Limpiar filtros
          </button>
        </div>
      </div>

      {/* ── Barra de selección múltiple ─────────────────────── */}
      {pedidosSeleccionados.length > 0 && (
      <div className={styles.seleccionBar}>
        <button className={styles.btnSeleccionar} onClick={todosSeleccionados ? onLimpiarSeleccion : onSeleccionarTodos}>
          {todosSeleccionados ? '☑' : '☐'} Seleccionar todos ({pedidosFiltrados.length})
        </button>
        <button className={styles.btnLimpiarSel} onClick={onLimpiarSeleccion}>✕ Limpiar</button>
        <span className={styles.cambiarEstadoLabel}>Cambiar estado a:</span>
        {ESTADO_FLOW.map((e) => (
          <button key={e} className={`${styles.btnEstadoMasivo} ${ESTADOS_BADGE[e] || ''}`} onClick={() => onCambiarEstadoMasivo(e)}>
            {e}
          </button>
        ))}
      </div>
      )}

      {/* ── Cards ───────────────────────────────────────────── */}
      {loading ? (
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.large} label="Cargando pedidos..." />
        </div>
      ) : pedidosFiltrados.length === 0 ? (
        <div className={styles.emptyState}>
          <Text>No hay pedidos que coincidan con los filtros aplicados.</Text>
        </div>
      ) : (
        <div className={styles.cardGrid}>
          {pedidosFiltrados.map((pedido) => {
            const items = getItemsForPedido(pedido.ID);
            const isLoadingItems = loadingItemsIds.includes(pedido.ID);
            const isSelected = pedidosSeleccionados.includes(pedido.ID);
            const isExpanded = expandedIds.includes(pedido.ID);
            const estadoClass = ESTADOS_BADGE[pedido.EstadoPedido] || styles.estadoNuevo;
            const isDelivery = pedido.MetodoEntrega === 'Delivery' || pedido.MetodoEntrega === 'Entrega a domicilio';
            const totalUnidades = items.reduce((s, i) => s + i.Cantidad, 0);
            // TODO: reemplazar por cálculo real desde campo de precio en SP
            const subtotalProductos = items.reduce((sum, item) => {
              const dollarIdx = item.Producto.lastIndexOf('$');
              if (dollarIdx < 0) return sum;
              const raw = item.Producto.slice(dollarIdx + 1).replace(/[.\s]/g, '').replace(',', '.');
              const precio = parseFloat(raw);
              return sum + (isNaN(precio) ? 0 : precio * item.Cantidad);
            }, 0);
            const total = subtotalProductos + (isDelivery ? costoEnvio : 0);

            return (
              <div
                key={pedido.ID}
                className={`${styles.pedidoCard} ${isSelected ? styles.pedidoCardSelected : ''}`}
                onClick={() => onVerDetalle(pedido)}
              >
                {/* ── Top: nro + checkbox + chevron ── */}
                <div className={styles.cardTopRow}>
                  <span className={styles.cardNro}>{pedido.Title}</span>
                  <div className={styles.cardTopActions}>
                    <div
                      className={styles.cardCheckArea}
                      onClick={(e) => { e.stopPropagation(); onToggleSeleccion(pedido.ID); }}
                    >
                      <Checkbox checked={isSelected} onChange={() => onToggleSeleccion(pedido.ID)} />
                    </div>
                    <span className={styles.cardChevron}>›</span>
                  </div>
                </div>

                {/* ── Estado badge ── */}
                <div className={styles.cardEstadoRow}>
                  <span className={`${styles.estadoBadge} ${estadoClass}`}>{pedido.EstadoPedido}</span>
                </div>

                {/* ── Productos box ── */}
                <div className={styles.cardProductosBox}>
                  <div className={styles.cardProductosHeader}>
                    <span className={styles.cardProductosHeaderTitle}>🎁 PEDIDO</span>
                    {!isLoadingItems && items.length > 0 && (
                      <span className={styles.cardProductosHeaderResumen}>
                        {items.length} prod. · {totalUnidades} u.
                      </span>
                    )}
                  </div>

                  {isLoadingItems ? (
                    <Spinner size={SpinnerSize.xSmall} />
                  ) : items.length === 0 ? (
                    <span className={styles.cardNoProductos}>Sin productos</span>
                  ) : (
                    items.slice().sort((a, b) => a.Orden - b.Orden).map((item) => {
                      const sepIdx = item.Producto.lastIndexOf(' – ');
                      const prodNombre = sepIdx >= 0 ? item.Producto.slice(0, sepIdx) : item.Producto;
                      const prodPrecio = sepIdx >= 0 ? item.Producto.slice(sepIdx + 3) : '';
                      return (
                        <div key={item.ID} className={styles.cardProductoItem}>
                          <span className={styles.cardProductoCantBadge}>{item.Cantidad}×</span>
                          <span className={styles.cardProductoNombre}>{prodNombre}</span>
                          {prodPrecio && <span className={styles.cardProductoPrecio}>{prodPrecio}</span>}
                        </div>
                      );
                    })
                  )}

                  {isDelivery && (
                    <div className={styles.cardEnvioRow}>
                      <span className={styles.cardEnvioIcon}>🚚</span>
                      <span className={styles.cardEnvioNombre}>Envío a domicilio</span>
                      <span className={styles.cardEnvioPrecio}>{formatMoney(costoEnvio)}</span>
                    </div>
                  )}
                </div>

                {/* ── Comentario ── */}
                {pedido.Comentarios && (
                  <div className={styles.cardComentario}>
                    📌 <em>{pedido.Comentarios}</em>
                  </div>
                )}

                {/* ── Cliente colapsable ── */}
                <div
                  className={styles.cardClienteToggle}
                  onClick={(e) => { e.stopPropagation(); onToggleExpand(pedido.ID); }}
                >
                  <span className={styles.cardClienteToggleArrow}>{isExpanded ? '▾' : '▸'}</span>
                  <span className={styles.cardClienteToggleNombre}>{pedido.NombreCompleto}</span>
                  {pedido.WhatsApp && (
                    <>
                      <span className={styles.cardClienteSep}>·</span>
                      <span className={styles.cardClientePhone}>📞 {pedido.WhatsApp}</span>
                    </>
                  )}
                  {!isExpanded && (
                    <span className={styles.cardClienteToggleLabel}>VER MÁS ›</span>
                  )}
                </div>

                {isExpanded && (
                  <div className={styles.cardClienteDetalle}>
                    <div className={styles.cardInfoRow}>
                      <span className={styles.cardInfoIcon}>📅</span>
                      <span className={styles.cardInfoText}>{formatFecha(pedido.FechaEntrega)}</span>
                      {pedido.HorarioAproximado && (
                        <>
                          <span className={styles.cardInfoIcon} style={{ marginLeft: 4 }}>⏱</span>
                          <span className={styles.cardInfoText}>{getFranja(pedido.HorarioAproximado)}</span>
                        </>
                      )}
                    </div>
                    <div className={styles.cardInfoRow}>
                      <span className={styles.cardInfoIcon}>{isDelivery ? '🚚' : '🏠'}</span>
                      <span className={styles.cardInfoText}>
                        {isDelivery ? (pedido.DireccionEntrega || 'Entrega a domicilio') : 'Retiro en local'}
                      </span>
                    </div>
                  </div>
                )}

                {/* ── Footer: total ── */}
                <div className={styles.cardFooter}>
                  <span className={styles.cardFooterLabel}>TOTAL A PAGAR</span>
                  <span className={`${styles.cardFooterMonto} ${total === 0 ? styles.totalCero : ''}`}>
                    {formatMoney(total)}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
