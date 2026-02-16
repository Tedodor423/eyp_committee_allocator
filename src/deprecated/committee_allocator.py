import pandas as pd
import re
from unicodedata import normalize

prefilled = True

def get_number(query=""):
    while not (output := input(query)).isnumeric(): print("toto není přirozené číslo")
    return int(output)

# load info
if prefilled: input_filename = "11th Regional Selection Conference of EYP CZ Hradec Králové 2022 Delegates All-in-One Form(1-70)_new.xlsx"
else:
    input_filename = input("Jméno vstupního Excel souboru VČETNĚ PŘÍPONY (může být i celá cesta): ")


if prefilled: committee_number, options_number, committee_column_range, name_column_range, school_column_letter = 8, 7, "AA,AL:AR", "F:G", "M"
else:
    name_column_range = input("Rozsah sloupců jména (ve formátu \"F:G\"): ")
    school_column_letter = input("Písmeno sloupce školy: ")
    committee_column_range = input("Rozsah sloupců výběru komisí (ve formátu \"AA,AL:AR\"): ")


# load file
try:
    name_column = pd.read_excel(input_filename, usecols=name_column_range)
    school_column = pd.read_excel(input_filename, usecols=school_column_letter)

    committee_columns = pd.read_excel(input_filename, usecols=committee_column_range)

except:
    print("Chyba při načítáni souboru - je správně jeho název a rozsah sloupců?")
    input()
    exit()

# parse names
names = []
for name in name_column.values.tolist():
    names.append(str(name[0]))
    for next_name in name[1:]:
        names[-1] += " " + str(next_name)
names.reverse()

# create main committee dict
committees = {}
for committee_name in list(committee_columns):
    committees[committee_name] = []
print("\n>> Načtena data od {0} delegátů\n>> Komise ({1}): {2}".format(len(committee_columns), len(committees), str(list(committees.keys()))[1:-1].translate({ord("'"): None})))

# parse schools
schools = list(school[0] for school in school_column.values.tolist())
schools.reverse()

# parse committee preferences
preferences_order = set()
committee_preferences = {}
for committee_name in committees:
    committee_preferences[committee_name] = list(reversed(committee_columns[committee_name].to_list()))

    # get committee preference order
    for preference in committee_preferences[committee_name]:
        preferences_order.add(preference)

# get committee size range with deviation:
committee_size = (len(committee_columns) // len(committees), -(len(committee_columns) // -len(committees)))
print(">> Průměrná velikost komisí: {0}-{1}".format(committee_size[0], committee_size[1]))

# committee_size_deviation = 0 if prefilled else get_number("Odchylka velikosti komisí: ")
# committee_size = (committee_size[0] - committee_size_deviation, committee_size[1] + committee_size_deviation)
# print(">> Upravená velikost komisí: {0}-{1}".format(committee_size[0], committee_size[1]))

# create dict for statistics
statistic = {}
for order in sorted(list(preferences_order)):
    statistic[order.replace(u'\xa0', u'')] = 0

# ASSIGN COMMITTEES
for highest_preference in sorted(list(preferences_order)):  # iterate through 1st to last preference order
    for committee_name in committees:

        # get all candidates with the highest preference
        candidates = []
        candidates_schools = []
        for person_index in range(len(committee_preferences[committee_name])):
            if committee_preferences[committee_name][person_index] == highest_preference:
                candidates.append(person_index)
                candidates_schools.append(schools[person_index])

        # shorten the list until there is right amount of delegates
        while len(candidates) > committee_size[1] - len(committees[committee_name]):
            # get most frequent schools
            schools_count = pd.Series(candidates_schools + list(c[1] for c in committees[committee_name])).value_counts()

            # remove last candidate (lists are reversed so pop the first) from the most frequent school
            for school in schools_count.index.to_list():  # find suitable candidate to pop
                if school in candidates_schools:
                    candidate_to_pop = school
                    break

            candidates.pop(candidates_schools.index(candidate_to_pop))
            candidates_schools.pop(candidates_schools.index(candidate_to_pop))

        # add candidates to committee
        for candidate_index in reversed(candidates):  # reversed to avoid eating our own tail when deleting in list
            committees[committee_name].append((names[candidate_index], schools[candidate_index], highest_preference.replace(u'\xa0', u'')))
            # remove candidate from input lists
            names.pop(candidate_index)
            schools.pop(candidate_index)
            for c_n in committees: committee_preferences[c_n].pop(candidate_index)

            # statistics
            statistic[highest_preference.replace(u'\xa0', u'')] += 1
            # if sorted(list(preferences_order)).index(highest_preference) > 3:
            #     print("Dotřídit: ", committees[committee_name][-1])

# for c_name in committees:
#     print(c_name)
#     for row in committees[c_name]:
#         print(row[0], row[1])
#     print("===============\n\n")

print(">> Satisfaction:", statistic)

# write to output

output_committees = []
for c_name in committees:
    output_committees += list([row[0], c_name, row[1], row[2]] for row in committees[c_name])

output_file = pd.DataFrame(output_committees, columns=["Jméno", "Komise", "Škola", "Preference"])

output_filename = input("Jméno výstupního Excel souboru (může být i celá cesta): ")
output_file.to_excel(output_filename, index=False)

print("\n>> Výsledky zapsány do souboru {0}".format(output_filename))
input()
