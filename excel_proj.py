import pandas as pd
from argparse import ArgumentParser
from pathlib import Path

if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument(
        'source_table',
        type=lambda p: Path(p).resolve().expanduser(),
        help='Path to excel table to retrieve data from'
    )
    parser.add_argument(
        'save_to_dir',
        type=lambda p: Path(p).resolve().expanduser(),
        help='Directory to save results table to',
    )
    args = parser.parse_args()
    df = pd.read_excel(args.source_table, index_col=0)

    # Stage 1
    def stage_1_formula(row, criteria_num: int):
        the_sum = sum(row[f'Criteria {i}'] for i in range(1, 6))
        return row['Ultimate weights of criteria, qi'] * row[f'Criteria {criteria_num}'] / the_sum


    for i in range(1, 6):
        df[f'Criteria {i}'] = df.apply(stage_1_formula, axis=1, criteria_num=i)

    # Stage 2
    name = 'The sums of weighted normalized maximizing indices of the windows, '

    results_plus = [df[df['*']][f'Criteria {i}'].sum() for i in range(1, 6)]
    results_minus = [df[~df['*']][f'Criteria {i}'].sum() for i in range(1, 6)]

    df.loc[name + 'S+j'] = [None, None, None, *results_plus]
    df.loc[name + 'S-j'] = [None, None, None, *results_minus]

    # Stage 3
    row_plus = df.loc[name + 'S+j']
    row_minus = df.loc[name + 'S-j']
    row_minus_sum = sum(row_minus[f'Criteria {i}'] for i in range(1, 6))


    def calculate_for_criteria(criteria_num: int):
        numerator = row_minus[f'Criteria {criteria_num}'] * row_minus_sum

        right_part_denom = sum(
            row_minus[f'Criteria {criteria_num}'] / row_minus[f'Criteria {i}']
            for i in range(1, 6)
        )
        denominator = row_minus[f'Criteria {criteria_num}'] * right_part_denom
        return row_plus[f'Criteria {criteria_num}'] + numerator / denominator


    significance = [calculate_for_criteria(i) for i in range(1, 6)]
    df.loc['Windows significance, Qj'] = [None, None, None, *significance]

    # Stage 4
    rating = []
    significance.sort(reverse=True)
    windows_significance_row = df.loc['Windows significance, Qj']

    for i in range(1, 6):
        value = windows_significance_row[f'Criteria {i}']
        rating.append(significance.index(value) + 1)

    df.loc['Priority of the alternative'] = [None, None, None, *rating]

    # Stage 5
    highest_rating_sig = windows_significance_row[f'Criteria {rating.index(1) + 1}']
    degree = []

    for i in range(1, 6):
        value = windows_significance_row[f'Criteria {i}']
        percent = (value / highest_rating_sig) * 100
        degree.append(percent)

    df.loc['Utility degree of the alternative (%)'] = [None, None, None, *degree]
    df.to_excel(args.save_to_dir / 'table_test.xlsx')
