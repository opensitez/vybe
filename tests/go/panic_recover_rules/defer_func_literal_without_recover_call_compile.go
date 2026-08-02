// vybe-test: go/panic_recover_rules/defer_func_literal_without_recover_call_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func run() { defer func() { _ = recover }
panic("x") }
func main() { run() }
