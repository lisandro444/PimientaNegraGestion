import * as React from 'react';
import { Text } from '@fluentui/react';
import type { IPedido } from '../../models/IPedido';
import styles from './PedidoManager.module.scss';

export interface IPedidoDetailViewProps {
  pedido: IPedido;
  savingEstado: boolean;
  onVolver: () => void;
  onCambiarEstado: (estado: string) => void;
  onEditar: () => void;
}

const ESTADOS_BADGE: Record<string, string> = {
  'Nuevo': styles.estadoNuevo,
  'Confirmado': styles.estadoConfirmado,
  'En preparación': styles.estadoEnPreparacion,
  'Listo': styles.estadoListo,
  'Entregado': styles.estadoEntregado,
  'Cancelado': styles.estadoCancelado
};

const FLUJO_SIGUIENTE: Record<string, string> = {
  'Nuevo': 'Confirmado',
  'Confirmado': 'Entregado'
};

const FLUJO_LABELS: Record<string, string> = {
  'Confirmado': '→ Confirmado',
  'Entregado': '→ Entregado'
};

function formatFechaLarga(isoDate?: string): string {
  if (!isoDate) return '-';
  const d = new Date(isoDate);
  return d.toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' });
}

function getFranja(horario: string): string {
  if (!horario) return '';
  const h = horario.toLowerCase();
  if (h.includes('12') || h.includes('13') || h.includes('14') || h.includes('mediod')) return 'Mediodía';
  if (h.includes('19') || h.includes('20') || h.includes('21') || h.includes('tarde')) return 'Tarde';
  if (h.includes('22') || h.includes('23') || h.includes('noche')) return 'Noche';
  return horario;
}

export const PedidoDetailView: React.FC<IPedidoDetailViewProps> = (props) => {
  const {
    pedido,
    savingEstado,
    onVolver, onCambiarEstado, onEditar
  } = props;

  const isDelivery = pedido.MetodoEntrega === 'Delivery' || pedido.MetodoEntrega === 'Entrega a domicilio';

  const estadoActual = pedido.EstadoPedido || 'Nuevo';
  const siguienteEstado = FLUJO_SIGUIENTE[estadoActual];
  const estaEntregado = estadoActual === 'Entregado';
  const estaCancelado = estadoActual === 'Cancelado';
  const waLink = pedido.WhatsApp
    ? `https://wa.me/${pedido.WhatsApp.replace(/\D/g, '')}`
    : undefined;

  return (
    <div className={styles.detailShell}>

      {/* ── Header ──────────────────────────────────────────── */}
      <div className={styles.detailHeader}>
        <div className={styles.detailHeaderLeft}>
          <button className={styles.btnVolver} onClick={onVolver}>← Volver</button>
          <div>
            <Text variant="xxLarge" block className={styles.pageTitle}>{pedido.Title}</Text>
            <Text variant="small" className={styles.pageSubtitle}>Detalle del pedido</Text>
          </div>
        </div>
        <div className={styles.detailHeaderRight}>
          <span className={`${styles.estadoBadge} ${ESTADOS_BADGE[estadoActual] || styles.estadoNuevo}`}>
            {estadoActual}
          </span>
          <button className={styles.btnVolver} onClick={onEditar}>✏ Editar</button>
        </div>
      </div>

      <div className={styles.detailBody}>
        <div className={styles.detailCardMain}>
          <Text variant="large" block className={styles.cardTitle}>Cliente y entrega</Text>

          <div className={styles.detailGrid}>
            <div className={styles.detailField}>
              <span className={styles.detailFieldLabel}>NOMBRE</span>
              <span className={styles.detailFieldValue}>{pedido.NombreCompleto}</span>
            </div>
            <div className={styles.detailField}>
              <span className={styles.detailFieldLabel}>WHATSAPP</span>
              <span className={styles.detailFieldValue}>
                📞 {pedido.WhatsApp || '-'}
                {waLink && (
                  <a
                    href={waLink}
                    target="_blank"
                    rel="noreferrer"
                    className={styles.waLink}
                  >
                    {' '}Abrir chat
                  </a>
                )}
              </span>
            </div>
            <div className={styles.detailField}>
              <span className={styles.detailFieldLabel}>FECHA DE ENTREGA</span>
              <span className={styles.detailFieldValue}>📅 {formatFechaLarga(pedido.FechaEntrega)}</span>
            </div>
            <div className={styles.detailField}>
              <span className={styles.detailFieldLabel}>FRANJA HORARIA</span>
              <span className={styles.detailFieldValue}>⏱ {getFranja(pedido.HorarioAproximado) || '-'}</span>
            </div>
            <div className={styles.detailField}>
              <span className={styles.detailFieldLabel}>MÉTODO DE ENTREGA</span>
              <span className={styles.detailFieldValue}>{pedido.MetodoEntrega || '-'}</span>
            </div>
            {isDelivery && pedido.DireccionEntrega && (
              <div className={styles.detailField}>
                <span className={styles.detailFieldLabel}>DIRECCIÓN</span>
                <span className={styles.detailFieldValue}>📍 {pedido.DireccionEntrega}</span>
              </div>
            )}
            <div className={styles.detailField}>
              <span className={styles.detailFieldLabel}>MÉTODO DE PAGO</span>
              <span className={styles.detailFieldValue}>{pedido.MetodoPago || '-'}</span>
            </div>
            {pedido.CubiertosDescartables && (
              <div className={styles.detailField}>
                <span className={styles.detailFieldLabel}>CUBIERTOS DESCARTABLES</span>
                <span className={styles.detailFieldValue}>Sí</span>
              </div>
            )}
          </div>

          {pedido.Comentarios && (
            <div className={styles.comentariosBox}>
              <span className={styles.comentariosLabel}>📌 NOTAS</span>
              <p className={styles.comentariosText}>{pedido.Comentarios}</p>
            </div>
          )}
          {estaEntregado && (
            <div className={styles.entregadoBadge}>✓ Pedido entregado</div>
          )}
        </div>

        <div className={styles.detailCardGestion}>
          <Text variant="large" block className={styles.cardTitle}>Gestión</Text>

          <div className={styles.gestionEstadoLabel}>ESTADO DEL PEDIDO</div>
          <div className={`${styles.estadoBadge} ${ESTADOS_BADGE[estadoActual] || styles.estadoNuevo}`} style={{ marginBottom: 12 }}>
            {estadoActual}
          </div>

          {!estaEntregado && !estaCancelado && (
            <>
              <div className={styles.gestionEstadoLabel}>Avanzar a:</div>
              <div className={styles.gestionBotones}>
                {siguienteEstado && (
                  <button
                    className={styles.btnAvanzarEstado}
                    onClick={() => onCambiarEstado(siguienteEstado)}
                    disabled={savingEstado}
                  >
                    {savingEstado ? '...' : FLUJO_LABELS[siguienteEstado]}
                  </button>
                )}
                <button
                  className={styles.btnCancelarEstado}
                  onClick={() => onCambiarEstado('Cancelado')}
                  disabled={savingEstado}
                >
                  ⊗ Cancelado
                </button>
              </div>
            </>
          )}

          {estaCancelado && (
            <button
              className={styles.btnAvanzarEstado}
              onClick={() => onCambiarEstado('Nuevo')}
              disabled={savingEstado}
            >
              ↺ Reactivar pedido
            </button>
          )}
        </div>
      </div>

    </div>
  );
};
