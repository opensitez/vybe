// vybe-test: go/control_flow_patterns_extra/switch_with_scoped_short_decl_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { switch value := 2; value { case 2: _ = value } }
