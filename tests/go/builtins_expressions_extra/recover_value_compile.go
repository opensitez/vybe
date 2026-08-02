// vybe-test: go/builtins_expressions_extra/recover_value_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func() { _ = recover() }()
panic("boom") }
