import { MSGraphClientV3 } from '@microsoft/sp-http';

// Fetch user events from Microsoft Graph API
export const fetchUserEvents = async (client: MSGraphClientV3): Promise<any[]> => {
  try {
    const response = await client.api('/me/events').get();
    // Map response to a format suitable for the calendar
    return response.value.map((event: any) => ({
      title: event.subject || 'No Title',
      start: event.start && event.start.dateTime ? new Date(event.start.dateTime) : new Date(),
      end: event.end && event.end.dateTime ? new Date(event.end.dateTime) : new Date()
    }));
  } catch (error) {
    console.error('Error fetching user events:', error);
    throw error;
  }
};

// Fetch transport locations from the TfL API
export const fetchLocations = async (): Promise<any[]> => {
  const pageSize = 25;
  const page = 1;
  try {
    const response = await fetch(`https://api.tfl.gov.uk/StopPoint/Mode/bus?page=${page}&pageSize=${pageSize}`);
    const data = await response.json();
    // Map response to dropdown options
    return data.stopPoints.map((stopPoint: any) => ({
      key: stopPoint.id,
      text: stopPoint.commonName
    }));
  } catch (error) {
    console.error('Error fetching locations:', error);
    throw error;
  }
};

// Fetch transport arrivals for a specific location
export const fetchArrivals = async (location: string): Promise<any[]> => {
  if (!location) return [];
  try {
    const response = await fetch(`https://api.tfl.gov.uk/StopPoint/${location}/Arrivals`);
    const data = await response.json();
    // Map response to a format suitable for the calendar
    return data.map((arrival: any) => ({
      title: `Bus ${arrival.lineName} to ${arrival.destinationName}`,
      start: new Date(arrival.expectedArrival),
      end: new Date(new Date(arrival.expectedArrival).getTime() + 15 * 60000)
    }));
  } catch (error) {
    console.error('Error fetching arrivals:', error);
    throw error;
  }
};