import pulp
import pandas as pd
import numpy as np

def optimize_project_schedule(tasks_df):
    prob = pulp.LpProblem("Project_Schedule_Optimization", pulp.LpMinimize)
    tasks_df['taskID'] = tasks_df['taskID'].astype(str)
    start_times = pulp.LpVariable.dicts("start_time",
                                      tasks_df.taskID.values,
                                      lowBound=0,
                                      cat='Continuous')
    
    # Compute PERT expected duration
    tasks_df['pert_duration'] = (tasks_df['bestCaseHours'] + 
                                4 * tasks_df['expectedHours'] + 
                                tasks_df['worstCaseHours']) / 6
    
    # Objective: Minimize project completion time
    total_project_duration = pulp.LpVariable("total_duration", lowBound=0)
    prob += total_project_duration
    
    # Add predecessor constraints and calculate task completion
    for idx, task in tasks_df.iterrows():
        task_id = str(task.taskID)
        duration = task['pert_duration']
        
        # Task completion time
        task_completion = start_times[task_id] + duration
        
        if pd.notna(task.predecessorTaskIDs) and task.predecessorTaskIDs:
            if isinstance(task.predecessorTaskIDs, str):
                pred_list = [p.strip() for p in task.predecessorTaskIDs.split(',') if p.strip()]
            else:
                pred_list = [str(task.predecessorTaskIDs).strip()]
            
            for pred in pred_list:
                try:
                    pred_duration = tasks_df.loc[tasks_df.taskID == pred, 'pert_duration'].iloc[0]
                    prob += start_times[task_id] >= start_times[pred] + pred_duration
                except IndexError:
                    print(f"Warning: Predecessor {pred} not found for task {task_id}")
                    continue
        
        prob += total_project_duration >= task_completion
    
    # Solve the optimization problem
    status = prob.solve()
    
    if status != 1:  # Not optimal
        raise ValueError(f"Could not find optimal solution. Status: {pulp.LpStatus[prob.status]}")
    
    # Prepare results
    results = {
        'status': pulp.LpStatus[prob.status],
        'total_project_duration': pulp.value(total_project_duration),
        'schedule': {},
        'task_details': []
    }
    
    # Extract detailed results
    for task_id in tasks_df.taskID:
        start_time = pulp.value(start_times[task_id])
        task_row = tasks_df.loc[tasks_df.taskID == task_id].iloc[0]
        
        task_result = {
            'taskID': task_id,
            'start_time': start_time,
            'pert_duration': task_row['pert_duration'],
            'best_case': task_row['bestCaseHours'],
            'expected_case': task_row['expectedHours'],
            'worst_case': task_row['worstCaseHours'],
            'end_time': start_time + task_row['pert_duration']
        }
        
        results['schedule'][task_id] = start_time
        results['task_details'].append(task_result)
    
    return results

def analyze_project_schedule(tasks_df):
    """
    Comprehensive project schedule analysis
    """
    # Run optimization
    optimization_results = optimize_project_schedule(tasks_df)
    
    # Convert results to DataFrame
    schedule_df = pd.DataFrame(optimization_results['task_details'])
    
    # Calculate additional metrics
    schedule_df['float_time'] = schedule_df['end_time'].max() - schedule_df['end_time']
    #schedule_df['criticality'] = schedule_df['float_time'].apply(lambda x: 'Critical' if abs(x) < 1e-10 else 'Non-Critical')
    
    # Sort by start time
    schedule_df = schedule_df.sort_values('start_time')
    
    return {
        'total_project_duration': optimization_results['total_project_duration'],
        'schedule': schedule_df,
        'optimization_status': optimization_results['status']
    }

def load_excel_data(file_path):
    """
    Load project schedule data from Excel
    """
    try:
        df = pd.read_excel(file_path)
        
        # Validate required columns
        required_columns = ['taskID', 'predecessorTaskIDs', 'bestCaseHours', 'expectedHours', 'worstCaseHours']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Convert taskID to string
        df['taskID'] = df['taskID'].astype(str)
        
        return df
        
    except Exception as e:
        raise Exception(f"Error loading Excel file: {str(e)}")

def main():
    file_path = 'project-plan-v003.xlsx'
    tasks_df = load_excel_data(file_path)
    project_analysis = analyze_project_schedule(tasks_df)
    
    print("\nProject Schedule Analysis:")
    print(f"Total Duration: {project_analysis['total_project_duration']:.2f} hours")
    print("\nDetailed Schedule:")
    print(project_analysis['schedule'].to_string())
    
    return project_analysis

if __name__ == "__main__":
    main()
