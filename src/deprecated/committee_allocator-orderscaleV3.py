"""

Changelog:
- order scale instead of selects
- best possible allocation - not enough

David Theodoor Nimrichtr 2023, EYP CZ
"""

from math import ceil, floor
import pandas as pd
import re
from progress.bar import Bar
import copy

prefilled = True

def get_number(query=""):
    while not (output := input(query)).isnumeric(): print("toto není přirozené číslo")
    return int(output)

# load info
if prefilled: input_filename = "KOL23 DELEGATES__BIBLE_committee_allocation.xlsx"
else:
    input_filename = input("Jméno vstupního Excel souboru VČETNĚ PŘÍPONY (může být i celá cesta): ")


if prefilled: 
    name_column_range = "A,C"
    committee_column_letter = "P"
    school_column_letter = "D"
    nationality_column_letter = "I"
    gender_column_letter = "G"

else:
    name_column_range = input("Rozsah sloupců jména (ve formátu \"A,C\"): ")
    committee_column_letter = input("Písmeno sloupce preferencí komisí: ")
    school_column_letter = input("Písmeno sloupce školy: ")
    nationality_column_letter = input("Písmeno sloupce národnosti: ")
    gender_column_letter = input("Písmeno sloupce genderu: ")


# load file
try:
    name_columns = pd.read_excel(input_filename, usecols=name_column_range)
    committee_pref_column = pd.read_excel(input_filename, usecols=committee_column_letter)
    school_column = pd.read_excel(input_filename, usecols=school_column_letter)
    gender_column = pd.read_excel(input_filename, usecols=gender_column_letter)
    nationality_column = pd.read_excel(input_filename, usecols=nationality_column_letter)

except:
    print("Chyba při načítáni souboru - je správně jeho název a rozsah sloupců?")
    input()
    exit()

# parse names
names = []
for name in name_columns.values.tolist():
    names.append(str(name[0]))
    for next_name in name[1:]:
        names[-1] += " " + str(next_name)
# names.reverse()

print("\n>> Načtena data od {0} delegátů".format(len(names)))

# create main committee dict - data from first row of committee pref column
committees = {}
for committee_name in committee_pref_column.values.tolist()[1][0].split(";"):
    committees[committee_name] = []

# parse committee preferences
committee_preferences = []
for committee_preference in committee_pref_column.values.tolist():
    committee_preferences.append(committee_preference[0].split(";")[:8])

commnum = len(committees)

print(" > Komise ({0}): {1}".format(commnum, str(list(committees.keys()))[1:-1]))

committee_size = (floor(len(names)/commnum), ceil(len(names)/commnum))

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

# pregenerate committees by preferences
preference_depth = get_number("\nDo kolikáté preference mám zohledňovat? (doporučeno 3): ")

precommittees = committees.copy()

for i in range(len(names)):
    for committee in committee_preferences[i][:preference_depth]:
        precommittees[committee].append(i)

print(">> Popularita komisí:")
for precommittee in precommittees.items():
    print(precommittee[0], ": ", len(precommittee[1]))

# SIMULATE
def populate_committees(committee_index, simulation_precommittees):
    for committee_iteration_index in range(len(simulation_precommittees[committee_index])-committee_size[1]): # <- here can be added variation of committee size
        if committee_index == commnum-2:
            bar.max += len(simulation_precommittees[committee_index])
        
        # copy new branch
        simulation_committees = copy.deepcopy(simulation_precommittees)

        # slice committee to iterated shape (optimalization)
        simulation_committees[committee_index] = simulation_committees[committee_index][committee_iteration_index:committee_iteration_index+committee_size[1]]
        
        # remove members from other committees
        smol = False
        for committee_member in simulation_committees[committee_index]:
            for following_committees_index in range(commnum-committee_index-1):
                try:
                    simulation_committees[committee_index+following_committees_index+1].remove(committee_member)
                    if len(simulation_committees[committee_index+following_committees_index+1]) <= committee_size[0]:
                        smol = True
                        break
                except ValueError:
                    pass
            if smol: break
        if smol:
            bar.next()
            continue
            
        if committee_index == commnum-1:
            # evaluate simulated committees
            print(simulation_committees)
        else:
            # pass generated and iterate with following comittees
            populate_committees(committee_index+1, simulation_committees)

        # iterate
        simulation_precommittees[committee_index].append(simulation_precommittees[committee_index][0])
        simulation_precommittees[committee_index].pop(0)


simulation_preprecommittees_names = list(committee_pref[0] for committee_pref in precommittees.items())
simulation_preprecommittees = list(committee_pref[1] for committee_pref in precommittees.items())

for simulation_committees_index in range(len(precommittees)):
    print(simulation_preprecommittees)
    with Bar('Simuluji {0} z {1}'.format(simulation_committees_index, commnum), max=1, fill='#', suffix='%(percent).2f%% - %(eta)ds') as bar:
        populate_committees(0, copy.deepcopy(simulation_preprecommittees))
    # move committees by one - start with a different one
    simulation_preprecommittees.append(simulation_preprecommittees[0])
    simulation_preprecommittees.pop(0)

