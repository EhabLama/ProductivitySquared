import { MSGraphClientV3 } from '@microsoft/sp-http';

export const fetchUserEvents = async (client: MSGraphClientV3): Promise<any[]> => {
  try {
    const response = await client.api('/me/events').get();
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

export const fetchLocations = async (): Promise<any[]> => {
  const pageSize = 25;
  const page = 1;
  try {
    const response = await fetch(`https://api.tfl.gov.uk/StopPoint/Mode/bus?page=${page}&pageSize=${pageSize}`);
    const data = await response.json();
    return data.stopPoints.map((stopPoint: any) => ({
      key: stopPoint.id,
      text: stopPoint.commonName
    }));
  } catch (error) {
    console.error('Error fetching locations:', error);
    throw error;
  }
};

export const fetchArrivals = async (location: string): Promise<any[]> => {
  if (!location) return [];
  try {
    const response = await fetch(`https://api.tfl.gov.uk/StopPoint/${location}/Arrivals`);
    const data = await response.json();
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