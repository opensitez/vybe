// vybe-test: go/function_literals_closures/mutual_recursive_closure_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
var a func(int) bool
var b func(int) bool
func init() { a = func(n int) bool { return b(n-1) }
b = func(n int) bool { return a(n-1) } }
func main() { _ = a(2) }
