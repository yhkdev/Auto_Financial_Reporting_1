

class Rule:

    def __init__(self, name, description):
        self.name = name
        self.description = description
        self.entries = []

    def add_entry(self, copyfrom, pasteto):
        self.entries.append([copyfrom, pasteto])

    def remove_entry(self, idx):
        del self.entries[idx]

    def edit_entry(self, idx, copyfrom, pasteto):
        self.entries[idx] = [copyfrom, pasteto]

    def update_name(self, input):
        self.name = input

    def update_description(self, input):
        self.description = input



# ----- Test -----
rule1 = Rule("Rule1", "Test Description")

# Test init()
assert rule1.name == "Rule1"
assert rule1.description == "Test Description"
assert rule1.entries == []

# Test update_name & update_description
rule1.update_name("RULE NAME TESTED")
rule1.update_description("DESCRIPTION TESTED")
assert rule1.name == "RULE NAME TESTED"
assert rule1.description == "DESCRIPTION TESTED"

# Test add_entry()
rule1.add_entry("B2", "B2")
rule1.add_entry("C2:C4", "C2:C4")
assert rule1.entries == [["B2", "B2"], ["C2:C4", "C2:C4"]]

# Test edit_entry()
rule1.edit_entry(0, "B3", "B3")
assert rule1.entries == [["B3", "B3"], ["C2:C4", "C2:C4"]]

# Test remove_entry()
rule1.remove_entry(0)
assert rule1.entries == [["C2:C4", "C2:C4"]]

