// vybe-test: go/defer_lifo_extended/defer_in_function_literal_label_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func main() { f := func() { defer func() {}() }
f() }
