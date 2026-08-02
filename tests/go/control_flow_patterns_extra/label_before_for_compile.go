// vybe-test: go/control_flow_patterns_extra/label_before_for_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { Start: for i := 0; i < 1; i++ { _ = i
break Start } }
