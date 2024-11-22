import * as React from 'react';
import styles from './SharepointDesign.module.scss';
import { ISharepointDesignProps } from './ISharepointDesignProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { Icon, PrimaryButton } from 'office-ui-fabric-react';

export interface ISharepointDesignState {

}

require("../assets/css/style.css");

export default class SharepointDesign extends React.Component<ISharepointDesignProps, ISharepointDesignState> {

  constructor(props: ISharepointDesignProps, state: ISharepointDesignState){
    super(props);

    this.state = {
      
    };
  }

  public render(): React.ReactElement<ISharepointDesignProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
     
        <div className="sharepointDesign">
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col">
                <div className='content'>
                  <div className="hero">
                    <div className="hero-content">
                      <h1 className="title">Microsoft SharePoint documentation</h1>
                      <p className='Description'>SharePoint documentation for IT professionals and admins</p>                      
                      {/* <p className="Description">Discover how to make the most of Power Automate with online training courses, docs, and videos covering product capabilities and how-to articles. Learn how to quickly create automated workflows between your favorite apps and services to synchronize files, get notifications, collect data, and more.</p> */}
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
