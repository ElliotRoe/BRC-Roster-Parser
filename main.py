from openpyxl import load_workbook


def abrev_last_name(name):
    name = name.strip()
    last_index = name.find(" ") + 1
    return name[0].upper() + name[1:last_index] + name[last_index].upper()


def get_challonge_name(sheet, row_num, team_column, member_column_1, member_column_2):
    return sheet[team_column + str(row_num)].value + " (" + abrev_last_name(
        sheet[member_column_1 + str(row_num)].value) + " & " + abrev_last_name(sheet[
                                                                                   member_column_2 + str(
                                                                                       row_num)].value) + ")"


if __name__ == '__main__':
    file_name = input("Please input the file path for the excel roster: ")
    column_team_name = input("Please input the team name column: ").strip().upper()
    column_name_1 = input("Please input the first member column: ").strip().upper()
    column_name_2 = input("Please input the second member column: ").strip().upper()
    workbook = load_workbook(filename=file_name, read_only=True)
    sheet = workbook.active

    print("\n-+-+-+-Names-+-+-+-\n")

    i = 2
    val = sheet["A" + str(i)].value
    while val:
        team_name = get_challonge_name(sheet, i, column_team_name, column_name_1, column_name_2)
        print(team_name)
        i = i + 1
        val = sheet["A" + str(i)].value
