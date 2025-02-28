import * as React from 'react';
import { Announced } from '@fluentui/react/lib/Announced';
import { TextField } from '@fluentui/react/lib/TextField';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode,
  DetailsRow,
  CheckboxVisibility,
  IDetailsListStyleProps,
  IDetailsListStyles,
} from '@fluentui/react/lib/DetailsList';
//import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyles } from '@fluentui/react/lib/Styling';
import IEmployeeDetails from '../../../models/IEmplyeeDetails';
import { IStyleFunctionOrObject } from '@fluentui/react';

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
  width: '50%',
});

const tableStyles: IStyleFunctionOrObject<
  IDetailsListStyleProps,
  IDetailsListStyles
> = {
  root: {
    display: 'flex',
  },
  contentWrapper: {
    root: {
      checkCell: {
        display: 'flex',
      },
    },
  },
  headerWrapper: {},
};

//Throwing error
// const textFieldStyles: Partial<ITextFieldStyles> = {
//   root: { maxWidth: '300px' },
// };

export interface IDetailsListBasicProps {
  items: IEmployeeDetails[];
  selectedItem: (item: number) => void;
}

export interface IDetailsListBasicState {
  listItems: IEmployeeDetails[];
  selectionDetails: string;
}

export default class DetailsListBasic extends React.Component<
  IDetailsListBasicProps,
  IDetailsListBasicState
> {
  private _selection: Selection;
  private _allItems: IEmployeeDetails[];
  private _columns: IColumn[];

  constructor(props: IDetailsListBasicProps, states: IDetailsListBasicState) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    // Populate with items for demos.
    this._allItems = this.props.items;

    this._columns = [
      {
        key: 'firstname',
        name: 'First Name',
        fieldName: 'FirstName',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'lastname',
        name: 'Last Name',
        fieldName: 'LastName',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'gender',
        name: 'Gender',
        fieldName: 'Gender',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'salary',
        name: 'Salary',
        fieldName: 'Salary',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
    ];

    this.state = {
      listItems: this.props.items,
      selectionDetails: this._getSelectionDetails(),
    };
  }

  componentDidUpdate(prevProps: IDetailsListBasicProps) {
    if (prevProps.items !== this.props.items) {
      this._allItems = this.props.items;
      this.setState({ listItems: this.props.items });
    }
  }

  public render() {
    const { listItems, selectionDetails } = this.state;

    return (
      <div>
        <div className={exampleChildClass}>{selectionDetails}</div>
        <Announced message={selectionDetails} />
        <TextField
          className={exampleChildClass}
          label="Filter by name:"
          onChange={this._onFilter}
          //styles={textFieldStyles}
        />
        <Announced
          message={`Number of items after filter applied: ${listItems.length}.`}
        />
        <DetailsList
          items={listItems}
          columns={this._columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selection={this._selection}
          selectionMode={SelectionMode.single}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="select row"
          styles={tableStyles}
          onRenderRow={(props: any) => (
            <DetailsRow
              {...props}
              onRenderCheck={() => this._onRenderCheckbox(props, props.item)}
              data-selection-disabled={true}
              onClick={(ev?: any) => {
                alert(props.item.Id);
              }}
            />
          )}
          //onRenderCheckbox={this._onRenderCheckbox}
          checkboxVisibility={CheckboxVisibility.always}
          onItemInvoked={this._onItemInvoked}
        />
      </div>
    );
  }

  private _onItemInvoked = (item: IEmployeeDetails): void => {
    alert(`Item invoked: ${item.Id}`);
  };

  private _onRenderCheckbox(props: any, item: IEmployeeDetails) {
    return <input type="checkbox" {...props} />;
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        this.props.selectedItem(0);
        return 'No items selected';
      case 1:
        //Send selected item id to parent component
        this.props.selectedItem(
          (this._selection.getSelection()[0] as IEmployeeDetails).Id
        );
        return (
          '1 item selected: ' +
          (this._selection.getSelection()[0] as IEmployeeDetails).Id
        );
      default:
        this.props.selectedItem(-1);
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    this.setState({
      listItems: text
        ? this._allItems.filter(
            (i) => i.FirstName.toLowerCase().indexOf(text) > -1
          )
        : this._allItems,
    });
  };
}
