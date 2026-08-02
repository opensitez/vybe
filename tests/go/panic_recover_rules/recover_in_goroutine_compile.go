// vybe-test: go/panic_recover_rules/recover_in_goroutine_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func main() { go func() { recover() }() }
