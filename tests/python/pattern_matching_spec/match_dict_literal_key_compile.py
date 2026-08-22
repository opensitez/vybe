# vybe-test: python/pattern_matching_spec/match_dict_literal_key_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

match {'kind': 'ok', 'value': 1}:
    case {'kind': 'ok', 'value': value}:
        print(value)
