// vybe-test: dart/control_flow_advanced/continue_in_while
// origin: languages/dart/tests/dart/test_control_flow_advanced.rs

void main() { var i = 0; while (i < 10) { i++; if (i % 2 == 0) continue; print(i); } }