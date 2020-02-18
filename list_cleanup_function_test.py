def name_checker(input_list, comparator_list):
    """Module compares list of generated_names with data in saved_projects and removes redundancies"""  
    for item in input_list:
        print("Processing item " + item + " from list1")
        if item in comparator_list:
            print("That's the stuff!")
            input_list.remove(item)
    return input_list

generated_names = ['a', 'b', 'c', 'd']
print("This is what list1 is like right now: ")
print(generated_names)
saved_names = ['a', 'c']
print("This is what list2 is like right now: ")
print(saved_names)
print("Passing list1 and list2 to name_checker for cleaning")
cleaned_list = name_checker(generated_names, saved_names)
print("Here comes your list as cleaned up by function")
print(cleaned_list)
