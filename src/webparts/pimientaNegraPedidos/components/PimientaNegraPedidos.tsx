import * as React from 'react';
import styles from './PimientaNegraPedidos.module.scss';
import type { IPimientaNegraPedidosProps } from './IPimientaNegraPedidosProps';
import PedidoManager from './pedidos/PedidoManager';

export default class PimientaNegraPedidos extends React.Component<IPimientaNegraPedidosProps> {
  public render(): React.ReactElement<IPimientaNegraPedidosProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.pimientaNegraPedidos} ${hasTeamsContext ? styles.teams : ''}`}>
        <PedidoManager {...this.props} />
      </section>
    );
  }
}
