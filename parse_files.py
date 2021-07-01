import glob
from argparse import ArgumentParser
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from warnings import warn

import pandas as pd
import xlsxwriter

DT_STRING = "%d-%m-%Y"
MAX_HOURS = 25


def read_file(filepath):
    with open(filepath, "rt") as f:
        found_date = False
        hour_counter = 0
        data = {
            "name": "",
            "date": None,
            "numbers": [],
            "signs": [],
            "unknown": None,
            "hours": None,
        }

        names = []
        for line in f:
            if not found_date:
                try:
                    data["date"] = datetime.strptime(line.strip(), DT_STRING)
                    found_date = True
                    data["name"] += "_".join(names)
                except ValueError:
                    names.append(line.strip())
            elif data["hours"] is None:
                data["hours"] = int(line.strip())
            elif hour_counter < data["hours"]:
                hour_counter += 1
                nums = line.strip().split(",")
                data["numbers"].append(float(nums[0]))
                data["signs"].append(nums[1])
            else:
                data["unknown"] = line.strip()
    return data


def parse_args():
    parser = ArgumentParser(description="This does... something.")
    parser.add_argument(
        "--input",
        "-i",
        type=str,
        default="**/*.dat",
        help="path to the files, as pattern: ex. `dir/OSDN_*.dat`",
    )
    parser.add_argument(
        "--out",
        "-o",
        type=Path,
        default="output.xlsx",
        help="Output file path of the tsv",
    )
    args, _ = parser.parse_known_args()
    return args


if __name__ == "__main__":
    args = parse_args()
    all_balances = defaultdict(list)
    for file_path in glob.glob(args.input, recursive=True):
        try:
            file_dict = read_file(file_path)
        except Exception:
            warn(f"Verify file {file_path} -- it is incorrect!!!")
            input('Press ENTER to close...')
        all_balances[file_dict["name"]].append(file_dict)
        if abs(file_dict["hours"] - 24) > 1:
            warn(f"File {file_path} has a wrong number of hours!!!")

    # Sorted raw data
    for balance_name, balance in all_balances.items():
        balance.sort(key=lambda x: x["date"])
        all_dates = [b["date"] for b in balance]
        if len(set(all_dates)) != len(all_dates):
            warn(f"Balance {balance_name} has repeating dates!!!")
        if len(set([len(b["numbers"]) for b in balance])) != 1:
            warn(
                f"Balance {balance_name} has different sizes of hours in some files!!!"
            )

    # Creating excel sheet
    with pd.ExcelWriter(args.out, engine="xlsxwriter") as writer:
        for i, (balance_name, balance) in enumerate(
            sorted(all_balances.items(), key=lambda x: x[0]), 1
        ):
            columns = list(range(1, MAX_HOURS + 1)) + ["sign"]
            rows = [d["date"].strftime(DT_STRING) for d in balance]
            data = []
            for balance_day in balance:
                # Raw numbers
                data_day = [n for n in balance_day["numbers"]][:MAX_HOURS]
                # Fix size 23 hours
                if len(data_day) == 23:
                    data_day = data_day[:2] + [None] + data_day[2:]
                # Fill to size
                data_day.extend([None] * (MAX_HOURS - len(data_day)))
                # Add signs (+ * -) to the end
                data_day.append(" ".join(sorted(set(balance_day["signs"]))))
                data.append(data_day)

            # creating single sheet
            df = pd.DataFrame(data, index=rows, columns=columns)
            sheetname = f"sheet_{i}"
            df.to_excel(writer, sheetname)
            worksheet = writer.sheets[sheetname]
            worksheet.write("A1", balance_name)

    input('Press ENTER to close...')
