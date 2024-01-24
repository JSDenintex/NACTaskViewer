export interface  ITaskAssignment {
    id: string;
    assignee: string;
    completedBy: string;
    completedDate: string;
    status: string;
    urls: {
      formUrl: string;
    };
  }
  
export interface  ITask {
    name: string;
    id: string;
    workflowName: string;
    status: string;
    createdDate: string;
    assigneeEmail: string;
    completedBy: string;
    dateCompleted: string | undefined;
    openTask: string;
    taskAssignments: ITaskAssignment[];
    outcomes: string[] | undefined;
    message: string;
}

export interface  IApiResponse {
    tasks: ITask[];
}