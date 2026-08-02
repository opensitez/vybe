// vybe-test: go/control_flow_patterns_extra/for_with_blank_condition_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { for ; true; { break } }
