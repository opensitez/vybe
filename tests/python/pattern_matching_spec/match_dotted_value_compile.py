# vybe-test: python/pattern_matching_spec/match_dotted_value_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

class Colors:
    RED = 'red'
color = 'red'
match color:
    case Colors.RED:
        pass
