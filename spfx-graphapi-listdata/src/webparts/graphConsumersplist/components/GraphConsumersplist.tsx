import * as React from 'react';
import styles from './GraphConsumersplist.module.scss';
import { IGraphConsumersplistProps } from './IGraphConsumersplistProps';
import { IGraphConsumersplistState } from './IGraphConsumersplistState';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from "@microsoft/sp-http";
import { PeoplePickerItemSuggestion } from 'office-ui-fabric-react';
import { IListItem } from './IListItem';

import {
  autobind,
  PrimaryButton,
  TextField,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode
} from 'office-ui-fabric-react';

// Configure the columns for the DetailsList component
let _listItemColumns = [
  {
    key: 'ContactPerson',
    name: 'Contact Person',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'ContactNumber',
    name: 'Contact Number',
    fieldName: 'ContactNumber',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'CompanyName',
    name: 'Company Name',
    fieldName: 'CompanyName',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'Country',
    name: 'Country',
    fieldName: 'Country',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
];


export default class GraphConsumersplist extends React.Component<IGraphConsumersplistProps, IGraphConsumersplistState> {

  public render(): React.ReactElement<IGraphConsumersplistProps> {
    return (
      <div className={ styles.graphConsumersplist }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>List Contact List Items</span>
          {
            (this.state.lists != null && this.state.lists.length > 0) ?
              <p className={ styles.form }>
              <DetailsList
                  items={ this.state.lists }
                  columns={ _listItemColumns }
                  setKey='set'
                  checkboxVisibility={ CheckboxVisibility.hidden }
                  selectionMode={ SelectionMode.none }
                  layoutMode={ DetailsListLayoutMode.fixedColumns }
                  compact={ true }
              />
            </p>
            : null
          }
</div>
          </div>
        </div>
      </div>
    );
  }



  constructor(props: IGraphConsumersplistProps, state: IGraphConsumersplistState) {
    super(props);

    // Initialize the state of the component
    this.state = {
      lists: []
    };
  }


  public componentDidMount(){
  // Log the current operation
  console.log("Using _searchWithGraph() method");

  this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {

      client
        .api("sites('root')/lists('Contactlist')/items?expand=fields")
        .version("v1.0")
        .get((err, res) => {

          if (err) {
            console.error(err);
            return;
          }

          // Prepare the output array
          var lists: Array<IListItem> = new Array<IListItem>();

          // Map the JSON response to the output array
          res.value.map((item: any) => {
            lists.push( {
              Title: item.fields.Title,
              ContactNumber: item.fields.ContactNumber,
              CompanyName: item.fields.CompanyName,
              Country: item.fields.Country,
            });
          });

          // Update the component state accordingly to the result
          this.setState(
            {
              lists: lists,
            }
          );
        });
    });


}

    /*var reactHandler = this;
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
    });*/
}
