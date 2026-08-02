# vybe-test: python/py_pattern_matching/test_py_match_walrus_equivalent_with_as
# origin: languages/python/tests/python/test_py_pattern_matching.rs

commands = [
    ["quit"],
    ["go", "north"],
    ["pick", "key", "rusty"],
    ["look"],
]

for cmd in commands:
    match cmd:
        case ["quit"]:
            print("Quitting")
        case ["go", direction]:
            print(f"Going {direction}")
        case ["pick", item, *adjectives]:
            print(f"Picking {' '.join(adjectives)} {item}")
        case [verb]:
            print(f"Unknown action: {verb}")
