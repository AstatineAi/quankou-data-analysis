import pandas as pd
import json
import re

def extract_question_id(column_name):
    match = re.match(r"(\d+)", str(column_name))
    if match:
        return match.group(1)
    return None

def main():
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    data = pd.read_excel(config['file_name'])
    data.dropna(subset=[data.columns[0]], inplace=True)
    results = pd.DataFrame()
    
    for idx, question in enumerate(config['questions'], start = 1):
        q_id = question['id']
        q_title = question['title']
        q_type = question['type']
        q_weight = question.get('weight', 0)
        
        question_cols = [col for col in data.columns if extract_question_id(col) == str(q_id)]
        
        if not question_cols:
            print(f"Question {q_id} not found in the data")
            continue
        
        if q_type == 'single_choice':
            q_range = question['range']
            data[question_cols[0]] = data[question_cols[0]].apply(lambda x: x if 0 <= x < q_range else q_range / 2)
            data[question_cols[0]] = data[question_cols[0]].apply(lambda x: (x + 1) / q_range)
            results[str(idx) + '. ' + q_title] = data[question_cols[0]] * q_weight
        elif q_type == 'multiple_choice':
            q_choices = question['choices']
            q_choices_weight = question['choices_weight']
            for weight, col in zip(q_choices_weight, question_cols):
                data[col] = data[col] * weight
            results[str(idx) + '. ' + q_title] = data[question_cols].sum(axis=1) * q_weight
        elif q_type == 'true_false':
            results[str(idx) + '. ' + q_title] = data[question_cols[0]].apply(lambda x: 1 - x) * q_weight
        elif q_type == 'info':
            results[str(idx) + '. ' + q_title] = data[question_cols[0]]
    
    numeric_columns = results.select_dtypes(include='number').columns
    results['总分'] = results[numeric_columns].sum(axis=1) * 100
    results = results.round(4)
    
    results.to_excel(config['output_file_name'] + '.xlsx', index=False)
    results.to_csv(config['output_file_name'] + '.csv', index=False)
    
    print("Analysis completed") 

if __name__ == '__main__':
    main()