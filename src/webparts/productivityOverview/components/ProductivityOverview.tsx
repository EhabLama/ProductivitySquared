// src/webparts/productivityOverview/components/ProductivityOverview.tsx

import * as React from 'react';
import styles from './ProductivityOverview.module.scss';
import type { IProductivityOverviewProps } from './IProductivityOverviewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { List } from '@fluentui/react/lib/List';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

interface IEmail {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
    };
  };
  bodyPreview: string;
}

interface IProductivityOverviewState {
  emails: IEmail[];
  loading: boolean;
}

export default class ProductivityOverview extends React.Component<IProductivityOverviewProps, IProductivityOverviewState> {
  constructor(props: IProductivityOverviewProps) {
    super(props);
    this.state = {
      emails: [],
      loading: true
    };
  }

  public componentDidMount(): void {
    this._fetchEmails();
  }

  private async _fetchEmails(): Promise<void> {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
      const response = await client
        .api('/me/messages')
        .version('v1.0')
        .select('subject,from,bodyPreview')
        .top(5)
        .get();

      this.setState({ emails: response.value, loading: false });
    } catch (error) {
      console.error('Error fetching emails', error);
      this.setState({ loading: false });
    }
  }

  public render(): React.ReactElement<IProductivityOverviewProps> {
    const { emails, loading } = this.state;
    const { userDisplayName } = this.props;

    return (
      <section className={styles.productivityOverview}>
        <div className={styles.welcome}>
          <h2>Welcome, {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <h3>Latest Emails</h3>
          {loading ? (
            <Spinner size={SpinnerSize.medium} label="Loading emails..." />
          ) : (
            <List
              items={emails}
              onRenderCell={(email: IEmail) => (
                <div key={email.id} className={styles.emailItem}>
                  <strong>{email.subject}</strong> from {email.from.emailAddress.name}
                  <p>{email.bodyPreview}</p>
                </div>
              )}
            />
          )}
        </div>
      </section>
    );
  }
}
