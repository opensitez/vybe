// vybe-test: go/control_flow_patterns_extra/select_empty_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func wait() { select {} }
func main() { _ = wait }
