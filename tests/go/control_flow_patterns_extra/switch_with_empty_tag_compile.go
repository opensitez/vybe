// vybe-test: go/control_flow_patterns_extra/switch_with_empty_tag_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { n := 2
switch { case n > 1: _ = n } }
