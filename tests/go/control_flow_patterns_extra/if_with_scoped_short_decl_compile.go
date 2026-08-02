// vybe-test: go/control_flow_patterns_extra/if_with_scoped_short_decl_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { if value := 1; value == 1 { _ = value } }
