import * as React from 'react';
import styles from './MicrosoftDesign.module.scss';
import { IMicrosoftDesignProps } from './IMicrosoftDesignProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

export interface IMicrosoftDesignState {

}

require("../assets/css/style.css");

export default class MicrosoftDesign extends React.Component<IMicrosoftDesignProps, IMicrosoftDesignState> {

  constructor(props: IMicrosoftDesignProps, state: IMicrosoftDesignState){
    super(props);

    this.state = {
      
    };
  }

  public render(): React.ReactElement<IMicrosoftDesignProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
        <div className="microsoftDesign">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col'>
                <div className="Microsoft-Header">
                  <div className='Microsoft'>
                    <div className="Microsoft-content">
                      <h2>Microsoft Ignite</h2>
                      <p className='Description'>Nov 19-22, 2024</p>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
    );
  }
}
