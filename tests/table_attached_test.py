from pysapscript.pysapscript import Sapscript


class TestRuns:
    def __init__(self):
        self.pss = Sapscript()
        self.window = self.pss.attach_window(0, 0)

    def test_runs(self):
        # get table
        table = self.window.read_shell_table("wnd[0]/usr/cntlGRID1/shellcont/shell")

        # print basic output
        print(str(table))
        print(repr(table))
        print(f"rows: {table.rows}, columns: {table.columns}")
        print(f"polars: {type(table.to_polars_dataframe())}, "
              f"pandas: {type(table.to_pandas_dataframe())}, "
              f"dict: {type(table.to_dict())}")
        print(f"column names: {table.get_column_names()}")

        # print to dict(s)
        print(table.to_dict())
        print(table.to_dicts())

        # slicing
        slci = table[5]
        print(slci)
        slc = table[2:4]
        print(slc)

        # iterating
        for row in table:
            print(row)

        # cell reading
        c = table.cell(1, 3)
        print(c)
        cs = table.cell(2, "SORTL")
        print(cs)

        # method
        table.select_rows([1, 3, 5])


if __name__ == "__main__":
    TestRuns().test_runs()
