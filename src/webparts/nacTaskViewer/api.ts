import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { ITask, IApiResponse } from './types';

// api.ts

export function getApiEndpoint(tenancyRegion: string): string {
  switch (tenancyRegion) {
    case 'us':
      return 'https://us.nintex.io';
    case 'eu':
      return 'https://eu.nintex.io';
    case 'au':
      return 'https://au.nintex.io';
    case 'ca':
      return 'https://ca.nintex.io';
    case 'uk':
      return 'https://uk.nintex.io';
    default:
      return 'https://us.nintex.io';
  }
}

export async function fetchTasks(
  baseUrl: string,
  queryParams: { [key: string]: string },
  accessToken: string
): Promise<ITask[]> {
  // Start with the base URL
  let url = `${baseUrl}/workflows/v2/tasks?`;
  
  // Construct the query parameters
  const queryString = buildQueryString(queryParams);
  
  // Append the unique cache-busting parameter
  const cacheBuster = `cacheBuster=${new Date().getTime()}`;
  url += queryString + (queryString ? '&' : '') + cacheBuster;
  
  console.log(`Generated Request URL: ${url}`);
  
  // Call getAccessToken to fetch the access token
  return getAccessToken(baseUrl, accessToken)
    .then((token: string) => {
      const requestHeaders: Headers = new Headers();
      requestHeaders.append('Authorization', `Bearer ${token}`);
      requestHeaders.append('Accept', 'application/json');
      const httpClientOptions: IHttpClientOptions = {
        headers: requestHeaders
      };
  
      return this.context.httpClient.get(url, HttpClient.configurations.v1, httpClientOptions);
    })
    .then((res: HttpClientResponse) => {
      if (!res.ok) {
          throw new Error(`API returned status: ${res.status}`);
      }
      return res.json();
    })
    .then((data: IApiResponse): ITask[] => { 
      if (!Array.isArray(data.tasks)) {
          throw new Error('Tasks is not an array or is missing from API response');
      }
      return data.tasks.map((task: ITask): ITask => ({
        name: task.name,
        workflowName: task.workflowName,
        id: task.id,
        status: task.status,
        createdDate: task.createdDate,
        assigneeEmail: task.taskAssignments[0]?.assignee || '',
        completedBy: task.taskAssignments[0]?.completedBy || '',
        dateCompleted: task.taskAssignments[0]?.completedDate || '',
        openTask: task.taskAssignments[0]?.urls?.formUrl || '',
        taskAssignments: task.taskAssignments,
        outcomes: task.outcomes || undefined,
        message: task.message,
      }));
    });
}

export async function getAccessToken(
  baseUrl: string,
  accessToken: string
): Promise<string> {
  const url = `${baseUrl}/authentication/v1/token`;
  const body = {
      client_id: accessToken,
      client_secret: accessToken,
      grant_type: 'client_credentials'
  };
  const headers: Headers = new Headers();
  headers.append('Content-Type', 'application/json');

  return this.context.httpClient.post(url, HttpClient.configurations.v1, {
      body: JSON.stringify(body),
      headers: headers
  })
  .then((res: HttpClientResponse) => {
      if (!res.ok) {
          throw new Error(`Server responded with status: ${res.status}`);
      }
      return res.json();
  })
  .then((data: { access_token: string }) => {
      if (!data.access_token) {
          throw new Error("No access token found in response");
      }
      return data.access_token;
  })
  .catch((err: Error) => {
    console.error("Error in getAccessToken:", err);
    throw err;
});
}

export async function confirmTaskCompletion(
  baseUrl: string,
  taskId: string,
  outcome: string,
  assignmentId: string,
  accessToken: string
): Promise<void> {
  const url = `${baseUrl}/workflows/v2/tasks/${taskId}/assignments/${assignmentId}`;
  const options = {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json, application/problem+json',
      Authorization: `Bearer ${accessToken}`
    },
    body: JSON.stringify({ outcome: outcome })
  };

  return fetch(url, options)
    .then(response => response.json())
    .then(data => {
      console.log('Task completed:', data);
      // Close the modal and possibly update the UI
      const modal = document.getElementById('confirmationModal');
      if (modal) {
        modal.style.display = 'none';
        this.fetchTasks();
      }
      // Additional logic to update UI
    })
    .catch(error => console.error('Error completing task:', error));
}

function buildQueryString(queryParams: { [key: string]: string }): string {
  return Object.entries(queryParams)
    .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
    .join('&');
}
