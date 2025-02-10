import {
  BaseButton,
  Button,
  CheckboxVisibility,
  DetailsList,
  DetailsListLayoutMode,
  PrimaryButton,
  SelectionMode,
  TextField,
} from '@fluentui/react';
import * as React from 'react';
import styles from './GraphConsumer.module.scss';
import { AadHttpClient, MSGraphClientV3 } from '@microsoft/sp-http';

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ClientMode } from '../../../common/ClientMode';

export interface IUserItem {
  displayName: string;
  mail: string;
  userPrincipalName: string;
}

export interface IGraphConsumerState {
  users: Array<IUserItem>;
  searchFor: string;
  messages: [
    {
      subject: string;
    }
  ];
}

export interface IGraphConsumerProps {
  clientMode: ClientMode;
  context: WebPartContext;
}
export default class GraphConsumer extends React.Component<
  IGraphConsumerProps,
  IGraphConsumerState
> {
  // Configure the columns for the DetailsList component
  private _usersListColumns = [
    {
      key: 'displayName',
      name: 'Display name',
      fieldName: 'displayName',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: 'mail',
      name: 'Mail',
      fieldName: 'mail',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: 'userPrincipalName',
      name: 'User Principal Name',
      fieldName: 'userPrincipalName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
  ];

  private _clientMode: any;
  //@ts-ignore
  private _context: any;

  constructor(props: IGraphConsumerProps, state: IGraphConsumerState) {
    super(props);

    this._clientMode = this.props.clientMode;
    this._context = this.props.context;

    // Initialize the state of the component
    this.state = {
      users: [],
      searchFor: '',
      messages: [
        {
          subject: '',
        },
      ],
    };

    this._onSearchForChanged = this._onSearchForChanged.bind(this);
    this._searchWithGraph = this._searchWithGraph.bind(this);
    this._searchWithAad = this._searchWithAad.bind(this);
  }

  private _onSearchForChanged = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue!, //It tells TypeScript that even though something looks like it could be null, it can trust you that it's not:
    });
  };

  private _getSearchForErrorMessage = (value: string): string => {
    // The search for text cannot contain spaces
    return value == null || value.length == 0 || value.indexOf(' ') < 0
      ? ''
      : `Error while searching`;
  };

  private _search = (
    event: React.MouseEvent<
      | HTMLAnchorElement
      | HTMLButtonElement
      | HTMLDivElement
      | BaseButton
      | Button,
      MouseEvent
    >
  ): void => {
    console.log(this._clientMode);

    // Based on the clientMode value search users
    switch (this._clientMode) {
      case ClientMode.aad:
        this._searchWithAad();
        break;
      case ClientMode.graph:
        this._searchWithGraph();
        break;
    }
  };

  public render(): React.ReactElement<IGraphConsumerProps> {
    return (
      <div className={styles.graphConsumer}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Search for a user!</span>
              <p className={styles.form}>
                <TextField
                  label={'Search'}
                  required={true}
                  onChange={this._onSearchForChanged}
                  onGetErrorMessage={this._getSearchForErrorMessage}
                  value={this.state.searchFor}
                />
              </p>
              <p className={styles.form}>
                <PrimaryButton
                  text="Search"
                  title="Search"
                  onClick={this._search}
                />

                <PrimaryButton
                  text="Get Emails"
                  title="Get Emails"
                  onClick={() => this.getDriveItems()}
                />
              </p>
              {this.state.users != null && this.state.users.length > 0 ? (
                <p className={styles.form}>
                  <DetailsList
                    items={this.state.users}
                    columns={this._usersListColumns}
                    setKey="set"
                    checkboxVisibility={CheckboxVisibility.hidden}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    compact={true}
                  />
                </p>
              ) : null}
              <p className={styles.form}>
                {this.state.messages.map((m) => (
                  <>
                    <span>{m.subject}</span>
                    <br />
                  </>
                ))}
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _searchWithGraph = (): void => {
    // Log the current operation
    console.log('Using _searchWithGraph() method');

    this.props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api('users')
          .version('v1.0')
          .select('displayName,mail,userPrincipalName')
          .filter(
            `(givenName eq '${escape(
              this.state.searchFor
            )}') or (surname eq '${escape(
              this.state.searchFor
            )}') or (displayName eq '${escape(this.state.searchFor)}')`
          )
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }

            // Prepare the output array
            var users: Array<IUserItem> = new Array<IUserItem>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push({
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });

            // Update the component state accordingly to the result
            this.setState({
              users: users,
            });
          });
      });
  };

  private _searchWithAad = (): void => {
    // Log the current operation
    console.log('Using _searchWithAad() method');

    // Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client.get(
          `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName&$filter=(givenName%20eq%20'${escape(
            this.state.searchFor
          )}')%20or%20(surname%20eq%20'${escape(
            this.state.searchFor
          )}')%20or%20(displayName%20eq%20'${escape(this.state.searchFor)}')`,
          AadHttpClient.configurations.v1
        );
      })
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        // Prepare the output array
        var users: Array<IUserItem> = new Array<IUserItem>();

        // Log the result in the console for testing purposes
        console.log(json);

        // Map the JSON response to the output array
        json.value.map((item: any) => {
          users.push({
            displayName: item.displayName,
            mail: item.mail,
            userPrincipalName: item.userPrincipalName,
          });
        });

        // Update the component state accordingly to the result
        this.setState({
          users: users,
        });
      })
      .catch((error) => {
        console.error(error);
      });
  };

  private getDriveItems() {
    let getMessages: string = 'me/messages';
    if (!this.props.context.msGraphClientFactory) {
      return;
    }
    this.props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        client
          .api(getMessages)
          .version('v1.0')
          .select('subject,sentDateTime,webLink')
          .top(5)
          .get((err: any, res: any): void => {
            if (err) {
              console.log('Getting error in retrieving mesages =>', err);
            }
            if (res) {
              console.log('Success');
              if (res && res.value.length) {
                console.log(res.value);
                this.setState({
                  messages: res.value,
                });
              }
            }
          });
      });
  }
}
