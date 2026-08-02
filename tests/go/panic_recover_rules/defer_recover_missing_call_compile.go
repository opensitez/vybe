// vybe-test: go/panic_recover_rules/defer_recover_missing_call_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func run() { defer recover()
panic("x") }
func main() { run() }
