"""
David Theodoor Nimrichtr, Adam Thomas Mezera 2026, EYP CZ
"""

from math import ceil, floor
import pandas as pd
from progress.bar import Bar
import tkinter as tk
from tkinter import filedialog
import os
from operator import itemgetter 

root = tk.Tk()
root.withdraw()    # hide main window
root.update()

def get_number(query=""):
    while not (output := input(query)).isnumeric(): print("toto není přirozené číslo")
    return int(output)

def get_float(query: str, default = None):
    query = query + " " if query.lstrip() == query else query
    output = ""
    while True:
        output = input(query + "(" + str(default) + ") " if isinstance(default, int | float) else query)
        if output == "" and isinstance(default, int | float): output = default; break 
        try:
            float(output)
        except:
            print("That is not a float")
            continue
        break
    return float(output)

last_session: list = []
if os.path.exists("./eyp_last_session"):
    while True:
        temp_input = input("Would you like to reload previous Excel file? (Yes = enter/No = N) ").lower().strip()
        if temp_input in ["n", "no", "ne"]:
            break
        elif temp_input in ["", "y", "yes", "ano", "jo"]:
            with open("./eyp_last_session", "r") as file:
                last_session = file.readlines()
            break

# load info
if last_session == []:
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
            parent=root,    # ensures dialog attaches to hidden root
            title="Select the Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    root.attributes("-topmost", False)
    root.destroy()    # close Tk properly
    
    with open("eyp_last_session", "w") as file:
        file.write(file_path + "\n")
else:
    file_path = last_session[0].strip()
    last_session.pop(0)


if len(last_session) == 0: 
    name_column_range = input("Rozsah sloupců jména (ve formátu \"A,C\"): ").upper()
    committee_column_letter = input("Písmeno sloupce preferencí komisí: ").upper()
    school_column_letter = input("Písmeno sloupce školy: ").upper()
    nationality_column_letter = input("Písmeno sloupce národnosti: ").upper()
    gender_column_letter = input("Písmeno sloupce genderu: ").upper()
    
    with open("eyp_last_session", "a") as file:
        file.write(name_column_range + "\n" + committee_column_letter + "\n" + school_column_letter + "\n" + nationality_column_letter + "\n" + gender_column_letter)
else:
    name_column_range = last_session[0].strip()
    committee_column_letter = last_session[1].strip()
    school_column_letter = last_session[2].strip()
    nationality_column_letter = last_session[3].strip()
    gender_column_letter = last_session[4].strip()

# load file
try:
    name_columns = pd.read_excel(file_path, usecols=name_column_range)
    committee_pref_column = pd.read_excel(file_path, usecols=committee_column_letter)
    school_column = pd.read_excel(file_path, usecols=school_column_letter)
    gender_column = pd.read_excel(file_path, usecols=gender_column_letter)
    nationality_column = pd.read_excel(file_path, usecols=nationality_column_letter)

except Exception as e:
    print("Chyba při načítáni souboru - je správně jeho název a rozsah sloupců?")
    print("Chyba: ", e)
    input()
    exit()

# parse names
names = []
for name in name_columns.values.tolist():
    names.append(str(name[0]))
    for next_name in name[1:]:
        names[-1] += " " + str(next_name)
DELNUM = len(names)

print("\n>> Načtena data {0} delegátů".format(DELNUM))

# create main committee dict - data from first row of committee pref column
committees = {}
for committee_name in sorted(committee_pref_column.values.tolist()[1][0].split(";")):
    committee_name = committee_name.strip()
    if committee_name == "":
        continue
    committees[committee_name] = []

# parse committee preferences
committee_preferences = []
for committee_preference in committee_pref_column.values.tolist():
    committee_preferences += [[]]
    for committee in committee_preference[0].split(";"):
        committee = committee.strip()
        if committee == "":
            continue
        committee_preferences[-1].append(committee)

COMNUM = len(committees)

print(" > Komise ({0}): {1}".format(COMNUM, str(list(committees.keys()))[1:-1]))

committee_size = (floor(DELNUM/COMNUM), ceil(DELNUM/COMNUM))

print(" > Velikost komisí: {0} - {1} delegátů".format(committee_size[0], committee_size[1]))

# parse schools
schoolset = set()
schools = []
for school in school_column.values.tolist():
    schoolset.add(school[0])
    schools.append(school[0])

print(" > Počet škol: {0}".format(len(schoolset)))

# parse nationalitiess
nationalityset = set()
nationalities = []
for nationality in nationality_column.values.tolist():
    nationalityset.add(nationality[0])
    nationalities.append(nationality[0])

print(" > Počet národností: {0}".format(len(nationalityset)))

# parse genders
genderset = set()
genders = []
for gender in gender_column.values.tolist():
    genderset.add(gender[0])
    genders.append(gender[0])

print(" > Počet genderů: {0}".format(len(genderset)))

# create prepopulated committees
preference_depth = get_number("\nDo kolikáté preference mám zohledňovat? (doporučeno 2-3): ")

for i in range(DELNUM):
    for committee in committee_preferences[i][:preference_depth]:
        committees[committee].append(i)


def print_committee_size():
    max_width = max(len(committee) for committee in committees.keys())
    for committee in sorted(committees.items(), key=lambda x: len(x[1]), reverse=True):
        print(f"{committee[0]:<{max_width}} : {len(committee[1])}")

print(">> Popularita komisí:")
print_committee_size()

# create dict for statistics
statistic = [0]*len(committees)


# DIVERSIFY COMMITTEES

print()
print("What weight would you like to give to:")
factors = ( # factor weights are applied for each other delegate with the same factor
    (genders, get_float("Gender?", 0.2)),
    (schools, get_float("Schools?", 0.25)), 
    (nationalities, get_float("Nationalities?", 0.2)),
)    # weight

preference_weight = get_float("Committee preference?", 0.2)

"""

Diversity maximalisation algorithm:

Evaluate all delegates in each committee
- for all diversity factors add nondiversity points equal to number of occurences of the value times weight of the factor
delete delegate with most nondiversity points
repeat until correct number of delegates
"""

# def check_delegate_deletion_nonviability(delegate_index, nondiverse_delegate_preindex):
#     # committee is too small
#     if len(committees[committee_order[int(nondiverse_delegate_preindex/DELNUM)]]) <= committee_size[0]:
#         return True

#     # delegate is in just one committee
#     delegate_presence = 0
#     for committee_name in committees:
#         if delegate_index in committees[committee_name]:
#             delegate_presence += 1

#     return delegate_presence <= 1    # True if cannot be deleted

# def old_algorithm():
#     with Bar('Maximalizuji diverzitu', max=DELNUM*preference_depth, fill='#', suffix='%(percent).2f%% - %(eta)ds') as bar:
#         for _deleted_delegate in range(DELNUM*preference_depth):
#             # calculate non diversity scores for all delegates per committee
#             nondiversity_scores = []
#             for committee_name in committees:
#                 nondiversity_scores += [0]*DELNUM
#                 committee_index = committee_order.index(committee_name)
#                 for factor in factors:
#                     factorscount = pd.Series(list(factor[0][delegate_index] for delegate_index in committees[committee_name]), dtype=str).value_counts()
#                     for delegate_index in committees[committee_name]:
#                         nondiversity_scores[DELNUM*committee_index+delegate_index] += factorscount[factor[0][delegate_index]]*factor[1]

#                 # factor in delegate preference
#                 for delegate_index in committees[committee_name]:
#                     nondiversity_scores[DELNUM*committee_index+delegate_index] += committee_preferences[delegate_index].index(committee_name)
            
#             # get delegate with highest nondiversity score
#             # - ensure delegate can be deleted

#             while check_delegate_deletion_nonviability(nondiverse_delegate_index := (nondiverse_delegate_preindex := nondiversity_scores.index(max(nondiversity_scores))) % DELNUM, nondiverse_delegate_preindex):
#                 nondiversity_scores[nondiverse_delegate_preindex] = -1

#                 if max(nondiversity_scores) == -1:
#                     break

#             # delete delegate
#             if max(nondiversity_scores) != -1:
#                 nondiverse_delegate_committee = committee_order[int(nondiverse_delegate_preindex/DELNUM)]
#                 try:
#                     committees[nondiverse_delegate_committee].remove(nondiverse_delegate_index)
#                 except ValueError:
#                     pass
#             bar.next()



def get_committee_diversity(committee_name:str) -> dict[int, float]: # calculate non diversity scores for all delegates per committee
    committee = committees[committee_name]

    nondiversity_scores: dict[int, float] = {}
    for factor in factors:
        factorscount = pd.Series(list(factor[0][delegate_index] for delegate_index in committee), dtype=str).value_counts()
        for delegate_index in committee:
            if delegate_index not in nondiversity_scores:
                nondiversity_scores[delegate_index] = .0
            nondiversity_scores[delegate_index] += factorscount[factor[0][delegate_index]]*factor[1]

    # factor in delegate preference - scaled with delegate amount
    for delegate_index in committee:
        nondiversity_scores[delegate_index] += committee_preferences[delegate_index].index(committee_name) * preference_weight * committee_size[1]/2
    
    return nondiversity_scores


def can_delete_delegate(delegate_index, committee_len) -> bool:
    if committee_len <= committee_size[0]:
        return False
    
    delegate_presence = 0
    for committee_name in committees:
        if delegate_index in committees[committee_name]:
            delegate_presence += 1
            if delegate_presence > 1: return True

    return False


def prune_committee(committee_name) -> bool: # returns true on successfully pruning a delegate
    committee_len = len(committees[committee_name])
    diversity = list(get_committee_diversity(committee_name).items())
    diversity.sort(key=itemgetter(1), reverse= True) # sorts by non diversity score

    for i in range(len(diversity)):
        if can_delete_delegate(diversity[i][0], committee_len):
            committees[committee_name].remove(diversity[i][0])
            return True
    return False


committee_order = list(committees)
print()


with Bar('Maximalizuji diverzitu', max=DELNUM*preference_depth-committee_size[0]*COMNUM, fill='#', suffix='%(percent).2f%% - %(eta)ds') as bar:
    not_pruned_committees = list(committees) # starts with all committees
    not_pruned_committees.sort(key=lambda x: len(committees[x]), reverse=True) # sorts them by amount of delegates
    biggest_committees = 1 # keeps track of how many committees have the same number of delegates

    while len(not_pruned_committees) > 0:
        if biggest_committees < len(not_pruned_committees):
            if len(committees[not_pruned_committees[biggest_committees]]) >= len(committees[not_pruned_committees[0]]):
                biggest_committees += 1
                continue
        
        for i in range(biggest_committees): # lists through and prunes only the biggest committees
            committee_name = not_pruned_committees[i]

            if len(committees[committee_name]) <= committee_size[0]: # makes sure not to prune too much (minimum committee size)
                not_pruned_committees.remove(committee_name)
                biggest_committees -= 1
                break
            
            if not prune_committee(committee_name): # if the committee can't be pruned anymore removes it
                not_pruned_committees.remove(committee_name)
                biggest_committees -= 1
                bar.next(len(committees[committee_name]) - committee_size[0]) # should make sure the bar is accurate, idk didn't really test
                break
            bar.next()

# old_algorithm()


print()
print(">> Velikost Komisí:")
print_committee_size()
print()
for del_i in range(DELNUM):
    committee_list = []
    for committee in committees:
        if del_i in committees[committee]:
            committee_list += [committee]

    if len(committee_list) == 0:
        raise ValueError(f"{names[del_i]} does not have a committee!")
    elif len(committee_list) > 1:
        print(f"{names[del_i]} je ve více komisí ({", ".join(committee_list)}) - rozřazení dokončete manuálně")

# output
output_committees = []
for committee_name in committees:
    output_committees += list([committee_name, names[del_i], schools[del_i], genders[del_i], nationalities[del_i], committee_preferences[del_i]] for del_i in committees[committee_name])
output_file = pd.DataFrame(output_committees, columns=["Committee", "Name", "School", "Gender", "Nationalities", "Preference"])

output_filename = input("\nJméno výstupního Excel souboru (může být i celá cesta): (output)")
if output_filename == "": output_filename = "output"
if not output_filename.endswith(".xlsx"): output_filename += ".xlsx"
output_file.to_excel(output_filename, index=False)

print("\n>> Výsledky zapsány do souboru {0}".format(output_filename))
input()

