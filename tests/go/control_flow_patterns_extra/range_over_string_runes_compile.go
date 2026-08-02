// vybe-test: go/control_flow_patterns_extra/range_over_string_runes_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { for _, r := range "hello" { _ = r } }
