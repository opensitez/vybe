// vybe-test: go/function_literals_closures/closure_capture_outer_in_defer_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { n := 1
defer func() { _ = n }() }
