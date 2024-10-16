import * as React from 'react';
import { ITransportArrivalsProps } from './ITransportArrivalsProps';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Calendar, momentLocalizer } from 'react-big-calendar';
import moment from 'moment';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import styles from './TransportArrivals.module.scss';
import { Text } from '@fluentui/react/lib/Text';
import { fetchUserEvents, fetchLocations, fetchArrivals } from '../../../api';

interface ITransportArrivalsState {
  selectedLocation: string;
  arrivals: any[];
  loading: boolean;
  calendarEvents: any[];
  locations: IDropdownOption[];
}

const localizer = momentLocalizer(moment);

export default class TransportArrivals extends React.Component<ITransportArrivalsProps, ITransportArrivalsState> {
  constructor(props: ITransportArrivalsProps) {
    super(props);
    this.state = {
      selectedLocation: '',
      arrivals: [],
      loading: false,
      calendarEvents: [],
      locations: []
    };
  }

  private fetchUserEvents = async (): Promise<void> => {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
      const events = await fetchUserEvents(client);
      this.setState((prevState) => ({
        calendarEvents: [...prevState.calendarEvents, ...events]
      }));
    } catch (error) {
      console.error('Error fetching user events:', error);
    }
  };

  private fetchLocations = async (): Promise<void> => {
    this.setState({ loading: true });
    try {
      const locations = await fetchLocations();
      this.setState({ locations, loading: false });
    } catch (error) {
      console.error('Error fetching locations:', error);
      this.setState({ loading: false });
    }
  };

  private fetchArrivals = async (location: string): Promise<void> => {
    this.setState({ loading: true, arrivals: [] });
    try {
      const transportEvents = await fetchArrivals(location);
      this.setState((prevState) => ({
        arrivals: transportEvents,
        loading: false,
        calendarEvents: [...prevState.calendarEvents, ...transportEvents]
      }));
    } catch (error) {
      console.error('Error fetching arrivals:', error);
      this.setState({ loading: false });
    }
  };

  private onLocationChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({ selectedLocation: option.key as string }, () => {
        this.fetchArrivals(this.state.selectedLocation);
      });
    }
  };

  public componentDidMount(): void {
    this.fetchUserEvents();
    this.fetchLocations();
  }

  public render(): React.ReactElement<ITransportArrivalsProps> {
    const { loading, selectedLocation, calendarEvents, locations } = this.state;

    return (
      <div className={styles.transportArrivals}>
        <Text variant="xxLarge" block styles={{ root: { fontWeight: 'bold' } }}>
          Select a London transport stop to add its arrivals to your calendar.
        </Text>
        {loading ? (
          <Spinner size={SpinnerSize.medium} label="Loading locations..." />
        ) : (
          <Dropdown
            className={styles.dropdown}
            label="Select Location"
            placeholder="Choose a transport stop to see upcoming arrivals"
            selectedKey={selectedLocation}
            onChange={this.onLocationChange}
            options={locations}
            calloutProps={{ calloutMaxHeight: 200 }}
          />
        )}
        {loading ? (
          <Spinner size={SpinnerSize.medium} label="Loading arrivals..." />
        ) : (
          <Calendar
            localizer={localizer}
            events={calendarEvents}
            startAccessor="start"
            endAccessor="end"
            style={{ height: 500 }}
          />
        )}
      </div>
    );
  }
}