# vybe-test: python/functions/function_with_defaults
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

def greet(name, greeting='hello'):
    print(greeting, name)
greet('world')
