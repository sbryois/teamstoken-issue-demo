import * as React from 'react';
import styles from './TeamsAuthIssueReproWp.module.scss';
import { ITeamsAuthIssueReproWpProps } from './ITeamsAuthIssueReproWpProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class TeamsAuthIssueReproWp extends React.Component<ITeamsAuthIssueReproWpProps, any> {

  public constructor(props, state) {
    super(props);

    this.state = {
      message: ""
    };

    this.onInit();
  }

  public async onInit(): Promise<void> {
    try {
      // URI of the AAD registered custom WebAPI being requested
      const appIdUri = 'https://<OUR-DISTANT-ENDPOINT>/api';
      // Just a helper to define whether an object is a Promise or not
      // Useful to conditionally resolve an exception which is a Promise
      const isPromise = (ex) => typeof ex === 'object' && typeof ex.then === 'function';

      this.props.context.aadTokenProviderFactory
        .getTokenProvider()
        .then((provider) => {
          provider
            .getToken(`${appIdUri}`, false)
            .then((tok) => {
              console.log('Provider/GetToken - Got a token', tok);
              this.setState({
                message: `Provider/GetToken - Got a token - ${tok}`
              });
            })
            .catch((e1) => {
              if (isPromise(e1)) {
                Promise.resolve(e1)
                  .then((s) => {
                    console.error('Provider/GetToken - Exception caught (Promise)', s['odata.error'].message.value);
                    this.setState({
                      message: `Provider/GetToken - Exception caught (Promise) - ${s['odata.error'].message.value}`
                    });
                  })
                  .catch((se) => 
                  {
                    console.error('Provider/GetToken - Exception caught (Promise)/Caught', se);
                    this.setState({
                      message: `Provider/GetToken - Exception caught (Promise)/Caught - ${se}`
                    });
                  });
              } else {
                console.error('Provider/GetToken - Exception caught (Normal)', e1);
                this.setState({
                  message: `Provider/GetToken - Exception caught (Normal) - ${e1}`
                });
              }
            });
        })
        .catch((e2) => {
          console.error('Get Token Provider went into catch', e2);
          this.setState({
            message: `Get Token Provider went into catch - ${e2}`
          });
          
        });
    } catch (e) {
      console.error('GENERAL EXCEPTION', e);
      this.setState({
        message: `Get Token GENERAL EXCEPTION - ${e}`
      });
    }
  }

  public render(): React.ReactElement<ITeamsAuthIssueReproWpProps> {
    return (
      <div className={ styles.teamsAuthIssueReproWp }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Teams Token Repro WP</span>
              <p className={ styles.subTitle }>Token Response below:</p>
              <p className={ styles.description }>{escape(this.state.message)}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
