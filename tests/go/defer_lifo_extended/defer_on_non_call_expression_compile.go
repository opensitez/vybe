// vybe-test: go/defer_lifo_extended/defer_on_non_call_expression_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func main() { defer 1 + 2 }
