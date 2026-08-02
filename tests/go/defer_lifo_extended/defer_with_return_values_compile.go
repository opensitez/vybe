// vybe-test: go/defer_lifo_extended/defer_with_return_values_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func main() { defer func() int { return 1 }() }
