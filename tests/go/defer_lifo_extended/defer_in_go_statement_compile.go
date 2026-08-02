// vybe-test: go/defer_lifo_extended/defer_in_go_statement_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func main() { defer func() { go func() {}() }() }
