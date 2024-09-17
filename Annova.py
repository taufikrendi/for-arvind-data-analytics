import pandas as pd
import numpy as np
from statsmodels.formula.api import ols
import statsmodels.api as sm

def load_data_from_excel(file_path, sheet_names):
    df_dict = pd.read_excel(file_path, sheet_name=sheet_names)
    return df_dict

def perform_anova(df, dependent_vars, independent_vars):
    anova_results = {}
    
    for dependent_var in dependent_vars:
        for independent_var in independent_vars:
            if not pd.api.types.is_categorical_dtype(df[independent_var]):
                df[independent_var] = df[independent_var].astype('category')
            
            model = ols(f'{dependent_var} ~ C({independent_var})', data=df).fit()
            anova_table = sm.stats.anova_lm(model, typ=2)
            
            anova_results[(dependent_var, independent_var)] = anova_table
    
    return anova_results


def main():
    file_path = '/Users/taufikrendi/document/anova/data.xlsx'
    sheet_names = ['cucitangan', 'APD', 'limbahinfeksius', 'linen', 'Pengetahuan', 'Sikap', 'Kewaspadaan']

    dependent_vars = ['Kewaspadaan']
    independent_vars = ['cucitangan', 'APD', 'limbahinfeksius', 'linen', 'Pengetahuan', 'sikap']

    data_dict = load_data_from_excel(file_path, sheet_names)
    
    for sheet_name, df in data_dict.items():
        print(f'Processing sheet: {sheet_name}')
        
        anova_results = perform_anova(df, dependent_vars, independent_vars)
        
        for (dep_var, indep_var), result in anova_results.items():
            print(f'ANOVA results for {dep_var} ~ {indep_var} in {sheet_name}')
            print(result)
            print()

if __name__ == '__main__':
    main()
