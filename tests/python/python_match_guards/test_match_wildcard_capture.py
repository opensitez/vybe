# vybe-test: python/python_match_guards/test_match_wildcard_capture
# origin: languages/python/tests/python/test_python_match_guards.rs

commands = ["start", "stop", "unknown"]
for cmd in commands:
    match cmd:
        case "start":
            print("starting")
        case "stop":
            print("stopping")
        case other:
            print(f"got: {other}")
