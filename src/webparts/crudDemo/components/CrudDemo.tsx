import * as React from 'react';
import styles from './CrudDemo.module.scss';
import type { ICrudDemoProps } from './ICrudDemoProps';
import { ISPListService } from '../../../services/core/spService/listServiceUsingPnP/ISPListServicePnP';
import { SPListService } from '../../../services/core/spService/listServiceUsingPnP/SPListServicePnP';
import { ISPLibraryService } from '../../../services/core/spService/libraryService/ISPLibraryServiceUsingPnP';
import { SPLibraryService } from '../../../services/core/spService/libraryService/SPLibraryServiceUsingPnP';
import {
  Button,
  FluentProvider,
  webLightTheme,
} from '@fluentui/react-components';
import { LibraryOption } from '../../../common/LibraryOption';
import IEmployeeDetails from '../../../models/IEmplyeeDetails';
import DetailsListBasic from '../customComponents/SimpleDetailsList';
import { ISPListServiceSPHTTP } from '../../../services/core/spService/listServiceUsingSPHttp/ISPListServiceSPHTTP';
import { SPListServiceSPHTTP } from '../../../services/core/spService/listServiceUsingSPHttp/SPListServiceSPHTTP';
import GraphConsumer from '../customComponents/GraphConsumer';
import * as $ from 'jquery';
import CustomFormModal from '../customComponents/CustomFormModal';
import Modal from '../../../controls/ModalPopUp';
import Loader from '../../../controls/SpinOverlay';

export interface ICrudDemoStates {
  libraryItems: any[];
  breadcrumbItems: any[];
  navigationLevel: number;
  listItems: IEmployeeDetails[];
  listItem: IEmployeeDetails;
  isModalOpen: boolean;
  showAlertModal: boolean;
  message: string;
  selectedItem: number;
  isUpdate: boolean;
  enableLoader: boolean;
}

export default class CrudDemo extends React.Component<
  ICrudDemoProps,
  ICrudDemoStates
> {
  //private _spListService: ISPListService;
  private _spLibraryService: ISPLibraryService;
  private _allItems: IEmployeeDetails[];
  constructor(props: ICrudDemoProps, states: ICrudDemoStates) {
    super(props);

    //this._spListService = new SPListService(this.props.context);
    this._spLibraryService = new SPLibraryService(this.props.context);
    this._allItems = [];

    this.state = {
      libraryItems: [],
      breadcrumbItems: [],
      navigationLevel: 0,
      listItems: this._allItems,
      listItem: {
        Id: 0,
        FirstName: '',
        LastName: '',
        Gender: '',
        Salary: '',
      },
      isModalOpen: false,
      showAlertModal: false,
      message: '',
      selectedItem: 0,
      isUpdate: false,
      enableLoader: false,
    };
  }

  public componentDidMount(): void {
    //const result = this._spListService.getListItems('EmployeeData');
    // result.then((r) => {
    //   console.log(r);
    // });
    console.log($('h4')[0].innerHTML);
    const rootFolderPath =
      this.props.context.pageContext.web.serverRelativeUrl + '/MyDocument';

    this.setState({
      breadcrumbItems: [
        {
          index: 0,
          Level: 'Home',
          RelativePath: rootFolderPath,
        },
      ],
    });
    this._getSPLibraryItems(rootFolderPath, false);
  }

  public _getSPLibraryItems = async (
    folderPath: string,
    isDifferentWeb: boolean
  ) => {
    let _getFolders: any[] = [],
      _getFiles: any[] = [];
    if (isDifferentWeb) {
      let siteUrl = this.props.context.pageContext.site.absoluteUrl.substring(
        0,
        this.props.context.pageContext.site.absoluteUrl.indexOf('site')
      );
      _getFolders =
        await this._spLibraryService.getFolderByServerRelativePathFromDifferentWeb(
          siteUrl,
          folderPath
        );
      _getFiles =
        await this._spLibraryService.getFilesByServerRelativePathFromDifferentWeb(
          siteUrl,
          folderPath
        );
    } else {
      _getFolders = await this._spLibraryService.getFoldersUsingRelativePath(
        folderPath
      );
      _getFiles = await this._spLibraryService.getFilesUsingRelativePath(
        folderPath
      );
    }

    Promise.all([_getFolders, _getFiles]).then((results) => {
      const [folderResult, fileResult] = results;

      console.log(folderResult);
      console.log(fileResult);

      let libObject = new Set();
      folderResult.forEach((r) => {
        libObject.add({
          Name: r.Name,
          ServerRelativeUrl: decodeURIComponent(r.ServerRelativeUrl),
          Folder: true,
          SourceUrl: '',
        });
      });
      fileResult.forEach((r) => {
        libObject.add({
          Name: r.Name,
          ServerRelativeUrl: decodeURIComponent(r.ServerRelativeUrl),
          Folder: r.Name.indexOf('.url') > -1 ? true : false,
          SourceUrl:
            r.Name.indexOf('.url') > -1
              ? decodeURIComponent(r.ListItemAllFields.SourceUrl)
              : '',
        });
      });

      this.setState({
        libraryItems: [...libObject],
      });
    });
  };

  public handleClick = (relativePath: string, FolderName: string) => {
    this.setState(
      {
        navigationLevel: this.state.navigationLevel + 1,
      },
      () => {
        this.setState({
          libraryItems: [],
          breadcrumbItems: [
            ...this.state.breadcrumbItems,
            {
              index: this.state.navigationLevel,
              Level: FolderName,
              RelativePath: relativePath,
            },
          ],
        });
      }
    );

    this._getSPLibraryItems(relativePath, false);
  };

  public handleClickOnFolderForUrl = (
    relativePath: string,
    FolderName: string
  ) => {
    console.log(relativePath);
    console.log(FolderName);
    this.setState(
      {
        navigationLevel: this.state.navigationLevel + 1,
      },
      () => {
        this.setState({
          libraryItems: [],
          breadcrumbItems: [
            ...this.state.breadcrumbItems,
            {
              index: this.state.navigationLevel,
              Level: FolderName,
              RelativePath: relativePath,
            },
          ],
        });
      }
    );
    this._getSPLibraryItems(relativePath, true);
  };

  public handleBreadcrumbClick = (
    index: number,
    menuLevel: string,
    RelativePath: string
  ) => {
    this.setState({
      breadcrumbItems: this.state.breadcrumbItems.filter(
        (x) => x.index <= index
      ),
      libraryItems: [],
    });
    this._getSPLibraryItems(RelativePath, false);
  };

  public render(): React.ReactElement<ICrudDemoProps> {
    return (
      <section className={styles.crudDemo}>
        {this.state.enableLoader ? <Loader /> : null}
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <h4>Selected Library: {this.props.selectedLibrary}</h4>
          <div className={styles.crudDemo__breadcrumb}>
            {this.state.breadcrumbItems.map((menu, i) => {
              return (
                <>
                  <span
                    tabIndex={menu.index}
                    onClick={() =>
                      this.handleBreadcrumbClick(
                        menu.index,
                        menu.Level,
                        menu.RelativePath
                      )
                    }
                  >
                    {menu.Level}
                  </span>
                  {this.state.breadcrumbItems.length - 1 > i ? (
                    <i className={styles.arrowRight}></i>
                  ) : (
                    ''
                  )}
                </>
              );
            })}
          </div>
          <div className={styles.tableContainer}>
            <table>
              <thead>
                <tr>
                  <th>
                    <td>Items</td>
                  </th>
                </tr>
              </thead>
              <tbody>
                {this.state.libraryItems.map((element) => {
                  return (
                    <tr>
                      <td>
                        {element.Folder ? (
                          element.Name.indexOf('.url') > -1 ? (
                            <a
                              onClick={() =>
                                this.handleClickOnFolderForUrl(
                                  element.SourceUrl,
                                  element.Name.split('.')[0]
                                )
                              }
                            >
                              {element.Name.split('.')[0]}
                            </a>
                          ) : (
                            <a
                              onClick={() =>
                                this.handleClick(
                                  element.ServerRelativeUrl,
                                  element.Name
                                )
                              }
                            >
                              {element.Name}
                            </a>
                          )
                        ) : (
                          element.Name
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        <div className={styles.CRUDOperationContainer}>
          <h3>CRUD Operation: {LibraryOption[this.props.selectedLibrary]}</h3>

          <FluentProvider theme={webLightTheme}>
            <div className={styles.formContainer}></div>
            <div className={styles.CRUDActionButtons}>
              <Button onClick={this._onLoadButtonClick}>Load Data</Button>
              <Button onClick={this._onCreateButtonClick}>Create</Button>
              <Button onClick={this._onUpdateButtonClick}>Update</Button>
              <Button>Delete</Button>
            </div>
            <div className={styles.tableContainer}>
              <DetailsListBasic
                items={this.state.listItems}
                selectedItem={this._getSelectionDetails}
              />
            </div>
          </FluentProvider>
        </div>
        <br />
        <FluentProvider>
          <div>
            <GraphConsumer
              clientMode={this.props.selectedLibrary}
              context={this.props.context}
            />
          </div>
        </FluentProvider>

        <CustomFormModal
          isOpen={this.state.isModalOpen}
          onClose={this._onCloseModal}
          onSave={this._onSubmitModal}
          onUpdate={this._onUpdateModal}
          item={this.state.listItem}
          isUpdate={this.state.isUpdate}
        />

        {this.state.showAlertModal && (
          <Modal
            message={this.state.message}
            onClose={() => this.setState({ showAlertModal: false })}
          />
        )}
      </section>
    );
  }

  private _onLoadButtonClick = () => {
    this.setState({ enableLoader: true }, () => {
      setTimeout(() => {
        this._loadListData();
      }, 1000);
    });
  };

  public _loadListData = async () => {
    if (this.props.selectedLibrary === LibraryOption['SP Service Using PnP']) {
      const _spListService: ISPListService = new SPListService(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      this.setState({ listItems: result, enableLoader: false });
    } else if (
      this.props.selectedLibrary === LibraryOption['SP Service Using SPHttp']
    ) {
      const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      this.setState({ listItems: result });
    } else if (
      this.props.selectedLibrary ===
      LibraryOption['SP Service Using MSGraph Client']
    ) {
      const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      console.log(result);
      this.setState({ listItems: result });
    } else {
    }
  };

  private _onCreateButtonClick = () => {
    this.setState({ isModalOpen: true, isUpdate: false });
  };

  private _onCloseModal = () => {
    this.setState({ isModalOpen: false });
  };

  //This method gets called when form submitted with details
  private _onSubmitModal = async (formData: IEmployeeDetails) => {
    if (this.props.selectedLibrary === LibraryOption['SP Service Using PnP']) {
      const _spListService: ISPListService = new SPListService(
        this.props.context
      );
      const result = await _spListService.createListItem(
        'EmployeeData',
        formData
      );

      if (result) {
        this.setState({
          showAlertModal: true,
          message: 'The record saved successfully!',
        });
      }
    } else if (
      this.props.selectedLibrary === LibraryOption['SP Service Using SPHttp']
    ) {
      const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      this.setState({ listItems: result });
    } else if (
      this.props.selectedLibrary ===
      LibraryOption['SP Service Using MSGraph Client']
    ) {
      const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      console.log(result);
      this.setState({ listItems: result });
    } else {
    }
  };

  private _onUpdateModal = async (formData: IEmployeeDetails) => {
    this.setState({ enableLoader: true });

    if (this.props.selectedLibrary === LibraryOption['SP Service Using PnP']) {
      const _spListService: ISPListService = new SPListService(
        this.props.context
      );
      const result = await _spListService.updateListItem(
        'EmployeeData',
        this.state.selectedItem,
        {
          Id: formData.Id,
          FirstName: formData.FirstName,
          LastName: formData.LastName,
          Gender: formData.Gender,
          Salary: formData.Salary,
        }
      );

      if (result) {
        this.setState({
          showAlertModal: true,
          message: 'The record updated successfully!',
          enableLoader: false,
        });
      }
    } else if (
      this.props.selectedLibrary === LibraryOption['SP Service Using SPHttp']
    ) {
      const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      this.setState({ listItems: result });
    } else if (
      this.props.selectedLibrary ===
      LibraryOption['SP Service Using MSGraph Client']
    ) {
      const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
        this.props.context
      );
      const result: IEmployeeDetails[] = await _spListService.getListItems(
        'EmployeeData'
      );
      console.log(result);
      this.setState({ listItems: result });
    } else {
    }
  };

  private _onUpdateButtonClick = async () => {
    if (this.state.selectedItem > 0) {
      if (
        this.props.selectedLibrary === LibraryOption['SP Service Using PnP']
      ) {
        const _spListService: ISPListService = new SPListService(
          this.props.context
        );
        const result = await _spListService.getListItemById(
          'EmployeeData',
          '',
          '',
          this.state.selectedItem
        );

        if (result) {
          this.setState({
            isModalOpen: true,
            isUpdate: true,
            listItem: result,
          });
        }
      } else if (
        this.props.selectedLibrary === LibraryOption['SP Service Using SPHttp']
      ) {
        const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
          this.props.context
        );
        const result: IEmployeeDetails[] = await _spListService.getListItems(
          'EmployeeData'
        );
        this.setState({ listItems: result });
      } else if (
        this.props.selectedLibrary ===
        LibraryOption['SP Service Using MSGraph Client']
      ) {
        const _spListService: ISPListServiceSPHTTP = new SPListServiceSPHTTP(
          this.props.context
        );
        const result: IEmployeeDetails[] = await _spListService.getListItems(
          'EmployeeData'
        );
        console.log(result);
        this.setState({ listItems: result });
      } else {
      }
    } else if (this.state.selectedItem < 0) {
      this.setState({
        showAlertModal: true,
        message: 'Please select only a single record from table to update!',
      });
    } else {
      this.setState({
        showAlertModal: true,
        message: 'Please select a record from table to update!',
      });
    }
  };

  private _getSelectionDetails = (item: number) => {
    this.setState({ selectedItem: item });
  };
}
