import pandas
import openpyxl
import argparse
import constant


def main(taken, out):

    df = pandas.read_excel(taken)
    cur_name = ''
    cur_dep = ''
    i = 0
    full_df = pandas.DataFrame()

    for name in df['Unnamed: 1']:
        if pandas.notna(name) and name != 'Итого:':
            if '(' in name:
                cur_dep = name
            elif constant.exception in name:
                cur_name = constant.exc_change
            else:
                cur_name = ''
                for n in name.split(' ')[-3:]:
                    cur_name += f"{n} "
        elif name == 'Итого:':
            datas = df.iloc[i]
            new_df = [data for data in datas]
            new_df.insert(1, name)
            new_df.insert(2, None)
            if new_df[0] != '':
                new_df = pandas.DataFrame([new_df])
                full_df = pandas.concat([full_df, new_df], ignore_index=True)

        else:
            datas = df.iloc[i]
            new_df = [data for data in datas]
            new_df.insert(1, cur_name)
            new_df.insert(2, cur_dep)
            new_df = pandas.DataFrame([new_df])

            if datas.iloc[0] == 1:
                empty = pandas.DataFrame([[None] * len(full_df.columns)], columns=full_df.columns)
                full_df = pandas.concat([full_df, empty], ignore_index=True)

            full_df = pandas.concat([full_df, new_df], ignore_index=True)

        i += 1

    #print(full_df)
    full_df = full_df.drop(full_df.columns[3], axis=1)
    full_df = full_df.drop(full_df.columns[4], axis=1)
    full_df = full_df.drop(full_df.columns[10], axis=1)
    full_df = full_df.drop(full_df.columns[10], axis=1)
    full_df = full_df.drop(full_df.columns[12], axis=1)
    full_df = full_df.drop(full_df.columns[15], axis=1)
    full_df = full_df.drop(full_df.columns[20], axis=1)
    full_df = full_df.drop(full_df.columns[24], axis=1)
    full_df = full_df.drop(full_df.columns[24], axis=1)

    df_screen = pandas.read_excel('from.xls', nrows=1)
    title_series = pandas.Series(df_screen.columns[0])

    with pandas.ExcelWriter(out, mode='w', engine='openpyxl') as writer:
        full_df.to_excel(writer, header=False, index=False, sheet_name="Отчёт")
    with pandas.ExcelWriter(out, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        title_series.to_excel(writer, header=False, index=False, sheet_name="Отчёт")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('taken', nargs='?', default=constant.taken)
    parser.add_argument('out', nargs='?', default=constant.out)

    args = parser.parse_args()

    print(args.taken, args.out)

    main(args.taken, args.out)