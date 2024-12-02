import * as React from 'react';
import styles from './SharepointDesign.module.scss';
import { ISharepointDesignProps } from './ISharepointDesignProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Item, sp } from '@pnp/sp/presets/all';
import { Icon, PrimaryButton } from 'office-ui-fabric-react';

export interface ISharepointDesignState {
  SharePointData : any;
}

require("../assets/css/style.css");

export default class SharepointDesign extends React.Component<ISharepointDesignProps, ISharepointDesignState> {

  constructor(props: ISharepointDesignProps, state: ISharepointDesignState){
    super(props);

    this.state = {
      SharePointData : ""
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
                      <h1 className="title">SharePoint Documentation</h1>
                      <p className='Description'>SharePoint documentation for IT professionals and admins</p>                      
                    </div>
                  </div>
                </div>

                <div className='ms-Grid-col'>
                  <div className='SharePoint-Info'>
                    {
                      this.state.SharePointData.length > 0 && 
                        this.state.SharePointData.map((item) => {
                          return(
                            <>
                              <div className='ms-Grid-col'>
                                <div className='sharepoint'>
                                  <div className='sharepoint-Icon'>
                                    <Icon iconName='SharepointAppIcon16' className='Sharepoint-ic' />
                                  </div>

                                  <div className='SharePoint-Details'>
                                    <a href='https://support.microsoft.com/en-us/office/what-is-sharepoint-97b915e6-651b-43b2-827d-fb25777f446f'>{item.Title}</a>
                                    <p>{item.Description}</p>
                                  </div>

                                </div>
                              </div>
                            </>
                          );
                        })
                    }
                  </div>  

                  {/* <div className="container">
                    <div className="card">
                      <div className="imgBx">
                        <img src="https://assets.codepen.io/4164355/shoes.png" />
                      </div>
                      <div className="contentBx">
                          <h2>Nike Shoes</h2>
                          <div className="size">
                            <h3>Size :</h3>
                            <span>7</span>
                            <span>8</span>
                            <span>9</span>
                            <span>10</span>
                          </div>
                          <div className="color">
                            <h3>Color :</h3>
                            <span></span>
                            <span></span>
                            <span></span>
                          </div>
                          <a href="#">Buy Now</a>
                      </div>
                    </div>
                  </div> */}


                </div>
              </div>
            </div>
          </div>
        </div>
    );
}

public async componentDidMount() {
  this.GetSharePointDetails();
}


  public async GetSharePointDetails() {
    const sharepoint = await sp.web.lists.getByTitle("SharePoint").items.select(
      "ID",
      "Title",
      "Description"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(sharepoint);

      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            Title : item.Title ? item.Title : "",
            Description : item.Description ? item.Description : "",
          });
        });
        this.setState({ SharePointData : AllData });
        console.log(this.state.SharePointData);
      }
    }).catch((Error) => {
      console.log("Error Retrived", Error);
    });
  }

}
