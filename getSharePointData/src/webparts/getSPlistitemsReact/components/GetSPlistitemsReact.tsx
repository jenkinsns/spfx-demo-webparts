import * as React from 'react';
import styles from './GetSPlistitemsReact.module.scss';
import { IGetSPlistitemsReactProps } from './IGetSPlistitemsReactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from 'jquery';

export interface IReactGetItemsState{
  items:[
        {
          "Title": "",
          "ContactNumber": "",
          "CompanyName":"",
          "Country":""
        }]
}

export default class GetSPlistitemsReact extends React.Component<IGetSPlistitemsReactProps, IReactGetItemsState> {

  public constructor(props: IGetSPlistitemsReactProps, state: IReactGetItemsState){
    super(props);
    this.state = {
      items: [
        {
          "Title": "",
          "ContactNumber": "",
          "CompanyName":"",
          "Country":""
        }
      ]
    };
  }

  public componentDidMount(){
    var reactHandler = this;
    jquery.ajax({
        url: `${this.props.currentsiteurl}/_api/web/lists/getbytitle('Contactlist')/items`,
        type: "GET",
        headers:{'Accept': 'application/json; odata=verbose;'},
        success: function(resultData) {
          reactHandler.setState({
            items: resultData.d.results
          });
        },
        error : function(jqXHR, textStatus, errorThrown) {
        }
    });
  }

  public render(): React.ReactElement<IGetSPlistitemsReactProps> {
    return (
      <div className={ styles.getSPlistitemsReact }>
      <div className={ styles.container }>
      <table className={styles.rtable}>
        <tr className={styles.rrow}>
            <th className={styles.rheader}>Contact Person</th>
            <th className={styles.rheader}>Contact Number</th>
            <th className={styles.rheader}>Company Name</th>
            <th className={styles.rheader}>Country</th>
        </tr>
        {this.state.items.map(function(item,key){
        return (
        <tr key={key}>
            <td className={styles.rcell}>{item.Title}</td>
            <td className={styles.rcell}>{item.ContactNumber}</td>
            <td className={styles.rcell}>{item.CompanyName}</td>
            <td className={styles.rcell}>{item.Country}</td>
        </tr>
        );
      })}
    </table>
    </div>
    </div>



    );
  }
}
