// vybe-test: go/panic_recover_rules/recover_used_as_statement_in_non_defer_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func main() { recover() }
