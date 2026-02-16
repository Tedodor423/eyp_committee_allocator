"""

Changelog:
- order scale instead of selects

David Theodoor Nimrichtr 2023, EYP CZ
"""

from math import ceil, floor
import pandas as pd
import re
from unicodedata import normalize
from progress.bar import Bar

prefilled = True

def get_number(query=""):
    while not (output := input(query)).isnumeric(): print("toto není přirozené číslo")
    return int(output)

# load info
if prefilled: input_filename = "committee allocatorV2/KOL23 DELEGATES__BIBLE_workshop_allocation.xlsx"
else:
    input_filename = "committee allocatorV2/KOL23 DELEGATES__BIBLE_committee_allocation.xlsx"  #input("Jméno vstupního Excel souboru VČETNĚ PŘÍPONY (může být i celá cesta): ")


if prefilled: 
    name_column_range = "A,C"
    committee_column_letter = "R"
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
delnum = len(names)

print("\n>> Načtena data {0} delegátů".format(delnum))

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

committee_size = (floor(delnum/commnum), ceil(delnum/commnum))

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

for i in range(delnum):
    for committee in committee_preferences[i][:preference_depth]:
        committees[committee].append(i)

print(committees)

print(">> Popularita komisí (1. a 2. preference):")
for committee in committees.items():
    print(committee[0], ": ", len(committee[1]))

# create dict for statistics
statistic = [0]*len(committees)


# DIVERSIFY COMMITTEES

factors = ((genders, 0.1), (schools, 0.2), (nationalities, 0.2))  # weight

"""

Diversity maximalisation algorithm:

Evaluate all delegates in each committee
- for all diversity factors add nondiversity points equal to number of occurences of the value times weight of the factor
delete delegate with most nondiversity points
repeat until correct number of delegates
"""

def check_delegate_deletion_nonviability(delegate_index):
    # committee is too small
    if len(committees[committee_order[int(nondiverse_delegate_preindex/delnum)]]) <= committee_size[0]:
        return True

    # delegate is in just one committee
    delegate_presence = 0
    for committee_name in committees:
        if delegate_index in committees[committee_name]:
            delegate_presence += 1

    return not delegate_presence-1  # True if cannot be deleted

committee_order = list(committees)
print()
with Bar('Maximalizuji diverzitu', max=delnum*preference_depth, fill='#', suffix='%(percent).2f%% - %(eta)ds') as bar:
    for _deleted_delegate in range(delnum*preference_depth):
        # calculate non diversity scores for all delegates per committee
        nondiversity_scores = []
        for committee_name in committees:
            nondiversity_scores += [0]*delnum
            committee_index = committee_order.index(committee_name)
            for factor in factors:
                factorscount = pd.Series(list(factor[0][delegate_index] for delegate_index in committees[committee_name]), dtype=str).value_counts()
                for delegate_index in committees[committee_name]:
                    nondiversity_scores[delnum*committee_index+delegate_index] += factorscount[factor[0][delegate_index]]*factor[1]

            # factor in delegate preference
            for delegate_index in committees[committee_name]:
                nondiversity_scores[delnum*committee_index+delegate_index] += committee_preferences[delegate_index].index(committee_name)
        
        # get delegate with highest nondiversity score
        # - ensure delegate can be deleted

        while check_delegate_deletion_nonviability(nondiverse_delegate_index := (nondiverse_delegate_preindex := nondiversity_scores.index(max(nondiversity_scores))) % delnum):
            nondiversity_scores[nondiverse_delegate_preindex] = -1

            if max(nondiversity_scores) == -1:
                break

        # delete delegate
        if max(nondiversity_scores) != -1:
            nondiverse_delegate_committee = committee_order[int(nondiverse_delegate_preindex/delnum)]
            try:
                committees[nondiverse_delegate_committee].remove(nondiverse_delegate_index)
            except ValueError:
                pass
        
        bar.next()

# immediate output
# for committee_name in committees:
#     print("======================================")
#     print(committee_name, len(committees[committee_name]))
#     for del_i in committees[committee_name]:
#         print(names[del_i], schools[del_i], genders[del_i], nationalities[del_i])


# output
output_committees = []
for committee_name in committees:
    output_committees += list([committee_name, names[del_i], schools[del_i], genders[del_i], nationalities[del_i], committee_preferences[del_i]] for del_i in committees[committee_name])
output_file = pd.DataFrame(output_committees, columns=["Committee", "Name", "School", "Gender", "Nationalities", "Preference"])

output_filename = input("\nJméno výstupního Excel souboru (může být i celá cesta): ")
if not output_filename.endswith(".xlsx"): output_filename += ".xlsx"
output_file.to_excel(output_filename, index=False)

print("\n>> Výsledky zapsány do souboru {0}".format(output_filename))
input()

# for preference_index in range(len()):  # iterate through 1st to last preference order
#     for committee_name in committees:

#         # get all candidates with the highest preference
#         candidates = []
#         candidates_schools = []
#         for person_index in range(len(committee_preferences[committee_name])):
#             if committee_preferences[committee_name][person_index] == highest_preference:
#                 candidates.append(person_index)
#                 candidates_schools.append(schools[person_index])

#         # shorten the list until there is right amount of delegates
#         while len(candidates) > committee_size[1] - len(committees[committee_name]):
#             # get most frequent schools
#             schools_count = pd.Series(candidates_schools + list(c[1] for c in committees[committee_name])).value_counts()

#             # remove last candidate (lists are reversed so pop the first) from the most frequent school
#             for school in schools_count.index.to_list():  # find suitable candidate to pop
#                 if school in candidates_schools:
#                     candidate_to_pop = school
#                     break

#             candidates.pop(candidates_schools.index(candidate_to_pop))
#             candidates_schools.pop(candidates_schools.index(candidate_to_pop))

#         # add candidates to committee
#         for candidate_index in reversed(candidates):  # reversed to avoid eating our own tail when deleting in list
#             committees[committee_name].append((names[candidate_index], schools[candidate_index], highest_preference.replace(u'\xa0', u'')))
#             # remove candidate from input lists
#             names.pop(candidate_index)
#             schools.pop(candidate_index)
#             for c_n in committees: committee_preferences[c_n].pop(candidate_index)

#             # statistics
#             statistic[highest_preference.replace(u'\xa0', u'')] += 1
#             # if sorted(list(preferences_order)).index(highest_preference) > 3:
#             #     print("Dotřídit: ", committees[committee_name][-1])

# # for c_name in committees:
# #     print(c_name)
# #     for row in committees[c_name]:
# #         print(row[0], row[1])
# #     print("===============\n\n")

# print(">> Satisfaction:", statistic)

# # write to output

# output_committees = []
# for c_name in committees:
#     output_committees += list([row[0], c_name, row[1], row[2]] for row in committees[c_name])

# output_file = pd.DataFrame(output_committees, columns=["Jméno", "Komise", "Škola", "Preference"])

# output_filename = input("Jméno výstupního Excel souboru (může být i celá cesta): ")
# output_file.to_excel(output_filename, index=False)

# print("\n>> Výsledky zapsány do souboru {0}".format(output_filename))
# input()
