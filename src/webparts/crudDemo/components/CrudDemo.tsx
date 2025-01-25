import * as React from 'react';
import styles from './CrudDemo.module.scss';
import type { ICrudDemoProps } from './ICrudDemoProps';
import { ISPListService } from '../../../services/core/spService/listServiceUsingPnP/ISPListServicePnP';
import { SPListService } from '../../../services/core/spService/listServiceUsingPnP/SPListServicePnP';

export default class CrudDemo extends React.Component<ICrudDemoProps, {}> {
  private _spListService: ISPListService;
  constructor(props: ICrudDemoProps) {
    super(props);

    this._spListService = new SPListService(this.props.context);
  }

  public componentDidMount(): void {
    const result = this._spListService.getListItems('EmployeeData');
    result.then((r) => {
      console.log(r);
    });
  }

  public render(): React.ReactElement<ICrudDemoProps> {
    return (
      <section className={`${styles.crudDemo}`}>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
        </div>
      </section>
    );
  }
}
