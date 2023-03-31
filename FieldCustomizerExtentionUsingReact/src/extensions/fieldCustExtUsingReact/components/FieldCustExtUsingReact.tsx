import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './FieldCustExtUsingReact.module.scss';

export interface IFieldCustExtUsingReactProps {
  text: string;
}

const LOG_SOURCE: string = 'FieldCustExtUsingReact';

export default class FieldCustExtUsingReact extends React.Component<IFieldCustExtUsingReactProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: FieldCustExtUsingReact mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: FieldCustExtUsingReact unmounted');
  }

  public render(): React.ReactElement<{}> {
   
    const mystyles = {
      color:'blue',
      width: `${this.props.text}px`,
      background: 'red'
    }

    return (
      <div className={styles.FieldCustExtUsingReact}>
        <div className={styles.cell}>
          <div style={mystyles}>
          { this.props.text }%
          </div>
        </div>      
      </div>
    );
  }
}
