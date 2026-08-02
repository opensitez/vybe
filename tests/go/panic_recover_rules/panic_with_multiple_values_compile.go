// vybe-test: go/panic_recover_rules/panic_with_multiple_values_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func main() { panic(1, 2) }
