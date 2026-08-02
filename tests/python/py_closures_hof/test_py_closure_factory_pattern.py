# vybe-test: python/py_closures_hof/test_py_closure_factory_pattern
# origin: languages/python/tests/python/test_py_closures_hof.rs

def validator_factory(**rules):
    def validate(data: dict) -> list:
        errors = []
        for field, rule in rules.items():
            if field not in data:
                errors.append(f"missing: {field}")
            elif not rule(data[field]):
                errors.append(f"invalid: {field}")
        return errors
    return validate

check = validator_factory(
    name=lambda v: isinstance(v, str) and len(v) > 0,
    age=lambda v: isinstance(v, int) and 0 < v < 150
)
print(check({"name": "Alice", "age": 30}))
print(check({"name": "", "age": 200}))
