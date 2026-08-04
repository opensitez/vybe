// vybe-test: dart/control_flow_advanced/nested_for
// origin: languages/dart/tests/dart/test_control_flow_advanced.rs

void main() { for (var i = 0; i < 3; i++) { for (var j = 0; j < 3; j++) { print('$i,$j'); } } }