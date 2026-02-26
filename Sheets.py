import string

print("Ibrahim Khan")

import_operators = ['+', '-', '/', '*']

class Spreadsheet:
    def __init__(self, rows, columns, name, cells_values={}):
        self.rows = max(10, rows)
        self.columns = max(10, columns)
        self.name = name
        self.cells_values = cells_values

    def set(self, position, value):
        # Check column/row bounds
        if int(string.ascii_lowercase.index(position[0].lower()) + 1) > self.columns:
            print("ERROR: Column Index out of reach")
            return
        if int(position[1:]) > self.rows:
            print("ERROR: Row Index out of reach")
            return

        # Direct integer
        if isinstance(value, int):
            self.cells_values[position] = value
            return value

        # Non-formula string
        if value[0] != '=':
            self.cells_values[position] = value
            return value

        # Formula
        formula = value[1:].replace(" ", "")
        curr_op = []
        for c in formula:
            if c in import_operators:
                curr_op.append(c)

        # Split formula into parts
        parts = []
        temp = ""
        for char in formula:
            if char in import_operators:
                if temp != "":
                    parts.append(temp)
                    temp = ""
            if char not in import_operators:
                temp += char
        if temp != "":
            parts.append(temp)

        # Evaluate each part (handle SUM and cell references)
        eval_parts = []
        for part in parts:
            if "SUM" in part:
                sum_range = part.replace("SUM", "").replace("[", "").replace("]", "")
                start, end = sum_range.split(":")
                if start[0] != end[0]:
                    ERROR = f"ERROR: {start[0]} and {end[0]} must be in same column"
                    self.cells_values[position] = ERROR
                    return ERROR
                s = 0
                for r in range(int(start[1:]), int(end[1:]) + 1):
                    key = start[0] + str(r)
                    if key in self.cells_values:
                        s += int(self.cells_values[key])
                    else:
                        ERROR = f"ERROR: Position {key} not defined in SUM range"
                        self.cells_values[position] = ERROR
                        return ERROR
                eval_parts.append(s)
            elif part.isdigit():
                eval_parts.append(int(part))
            else:  # single cell reference
                if part in self.cells_values:
                    eval_parts.append(int(self.cells_values[part]))
                else:
                    ERROR = f"ERROR: Position {part} doesn't exist"
                    self.cells_values[position] = ERROR
                    return ERROR

        # If only one part (SUM-only or single reference)
        if len(eval_parts) == 1:
            self.cells_values[position] = eval_parts[0]
            return eval_parts[0]

        # Apply operators
        result = eval_parts[0]
        for i, op in enumerate(curr_op):
            val = eval_parts[i+1]
            if op == '+':
                result += val
            elif op == '-':
                result -= val
            elif op == '*':
                result *= val
            elif op == '/':
                result /= val

        self.cells_values[position] = result
        return result

    def output_cell(self, position):
        if int(string.ascii_lowercase.index(position[0].lower()) + 1) > self.columns:
            print("ERROR: Column Index out of reach")
            return
        if int(position[1:]) > self.rows:
            print("ERROR: Row Index out of reach")
            return
        if position not in self.cells_values:
            print(f"ERROR: Position {position} doesn't exist")
            return
        return {position: self.cells_values[position]}


# --- TEST ---
sheet1 = Spreadsheet(5, 5, "sheet1")
sheet1.set("A1", 12)
sheet1.set("A5", 10)
sheet1.set("A6", 10)
sheet1.set("A7", 10)

# SUM-only
sheet1.set("B3", "=SUM[A5:A7]")  # 30
print(sheet1.output_cell("B3"))

# SUM + arithmetic
sheet1.set("B4", "=SUM[A5:A7]+A1")  # 42
print(sheet1.output_cell("B4"))

# SUM with multiple operators
sheet1.set("B5", "=SUM[A5:A7]*A1-A6")  # 30*12-10 = 350
print(sheet1.output_cell("B5"))
