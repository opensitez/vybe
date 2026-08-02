// vybe-test: go/higher_order_functions/mutual_recursive_functions
// origin: languages/go/tests/go/test_higher_order_functions.rs
// vybe-test-mode: compile

package main
var even func(int) bool
var odd func(int) bool
func init() { even = func(n int) bool { if n == 0 { return true }
return odd(n-1) }
odd = func(n int) bool { if n == 0 { return false }
return even(n-1) } }
func main() { _ = even(4) }
