import string

global import_operators
import_operators = ['+', '-', '/', '*']

class Spreadsheet:
    def __init__(self, rows, columns, name, cells_values={}):

        self.rows = max(5, rows)
        self.columns = max(5, columns)

        self.name = name
        self.cells_values = cells_values

    def set(self, position, value):

        if int(string.ascii_lowercase.index(position[0].lower()) + 1) > self.columns:
            print("ERROR: Column Index out of reach")

        elif int(position[1:]) > self.rows:
            print("ERROR: Row Index out of reach")

        else:

            if isinstance(value, int):
                self.cells_values[position] = value
                return value

            if value[0] != '=':
                self.cells_values[position] = value
                return value
            else:
                op_found = False
                curr_op = []
                for a in value:
                    for op in import_operators:
                        if op == a:
                            op_found = True
                            curr_op.append(op)
                if not op_found:
                    # NEW: check for single-cell reference
                    cleaned = value.replace(" ", "").replace("=", "")
                    if cleaned in self.cells_values:
                        self.cells_values[position] = self.cells_values[cleaned]
                        return self.cells_values[cleaned]
                    else:
                        ERROR = "ERROR: Dependencies require operator or valid cell reference"
                        return ERROR
                else:
                    cleaned = value.replace(" ", "").replace("=", "")


                    for op in import_operators:
                        cleaned = cleaned.replace(op, " ")

                    positions__ = cleaned.split()
                    print(len(positions__))

                    if all((pos in self.cells_values) for pos in positions__ if not pos.isdigit()):
                        sum = self.cells_values[positions__[0]]

                        for m in range(1, len(positions__)):
                            if curr_op[m - 1] == '+':
                                if positions__[m][0].isalpha():
                                    sum += self.cells_values[positions__[m]]
                                else:
                                    sum += int(positions__[m])
                            if curr_op[m - 1] == '-':
                                if positions__[m][0].isalpha():
                                    sum -= self.cells_values[positions__[m]]
                                else:
                                    sum -= int(positions__[m])

                            if curr_op[m - 1] == '/':
                                if positions__[m][0].isalpha():
                                    sum /= self.cells_values[positions__[m]]
                                else:
                                    sum /= int(positions__[m])

                            if curr_op[m - 1] == '*':
                                if positions__[m][0].isalpha():
                                    sum *= self.cells_values[positions__[m]]
                                else:
                                    sum *= int(positions__[m])

                    else:
                        ERROR = "ERROR: Some position argument cells have no assigned integer"
                        return ERROR

                self.cells_values[position] = sum
                return sum




    def output_cell(self, position):

        if int(string.ascii_lowercase.index(position[0].lower()) + 1) > self.columns:
            print("ERROR: Column Index out of reach")

        elif int(position[1:]) > self.rows:
            print("ERROR: Row Index out of reach")

        elif  position not in self.cells_values:
            print("ERROR: Position value doesn't exist")

        else:
            return {position: self.cells_values[position]}



sheet1 = Spreadsheet(5, 5, "sheet1")
sheet1.set("A1", 10)
sheet1.set("B2", 20)
sheet1.set("C3", 10)
sheet1.set("B3", "=A1 + C3 * 2")
print(sheet1.output_cell("B3"))
