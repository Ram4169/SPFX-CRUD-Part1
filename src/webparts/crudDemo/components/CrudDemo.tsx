import * as React from 'react';
import styles from './CrudDemo.module.scss';
import type { ICrudDemoProps } from './ICrudDemoProps';
import { ISPListService } from '../../../services/core/spService/listServiceUsingPnP/ISPListServicePnP';
import { SPListService } from '../../../services/core/spService/listServiceUsingPnP/SPListServicePnP';
import { ISPLibraryService } from '../../../services/core/spService/libraryService/ISPLibraryServiceUsingPnP';
import { SPLibraryService } from '../../../services/core/spService/libraryService/SPLibraryServiceUsingPnP';

export interface ICrudDemoStates {
  libraryItems: any[];
  breadcrumbItems: any[];
  navigationLevel: number;
}

export default class CrudDemo extends React.Component<
  ICrudDemoProps,
  ICrudDemoStates
> {
  private _spListService: ISPListService;
  private _spLibraryService: ISPLibraryService;
  constructor(props: ICrudDemoProps, states: ICrudDemoStates) {
    super(props);

    this._spListService = new SPListService(this.props.context);
    this._spLibraryService = new SPLibraryService(this.props.context);

    this.state = {
      libraryItems: [],
      breadcrumbItems: [],
      navigationLevel: 0,
    };
  }

  public componentDidMount(): void {
    const result = this._spListService.getListItems('EmployeeData');
    result.then((r) => {
      console.log(r);
    });

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
      folderResult.forEach((r) => {
        this.setState({
          libraryItems: [
            ...this.state.libraryItems,
            {
              Name: r.Name,
              ServerRelativeUrl: decodeURIComponent(r.ServerRelativeUrl),
              Folder: true,
              SourceUrl: '',
            },
          ],
        });
      });
      fileResult.forEach((r) => {
        this.setState({
          libraryItems: [
            ...this.state.libraryItems,
            {
              Name: r.Name,
              ServerRelativeUrl: decodeURIComponent(r.ServerRelativeUrl),
              Folder: r.Name.indexOf('.url') > -1 ? true : false,
              SourceUrl:
                r.Name.indexOf('.url') > -1
                  ? decodeURIComponent(r.ListItemAllFields.SourceUrl)
                  : '',
            },
          ],
        });
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
      </section>
    );
  }
}
