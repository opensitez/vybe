// vybe-test: dart/control_flow_advanced/break_in_for_in
// origin: languages/dart/tests/dart/test_control_flow_advanced.rs

void main() { for (var x in [1,2,3,4,5]) { if (x == 3) break; } }