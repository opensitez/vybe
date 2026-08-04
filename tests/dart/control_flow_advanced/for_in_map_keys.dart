// vybe-test: dart/control_flow_advanced/for_in_map_keys
// origin: languages/dart/tests/dart/test_control_flow_advanced.rs

void main() { var m = {'a': 1, 'b': 2}; for (var k in m.keys) { print(k); } }