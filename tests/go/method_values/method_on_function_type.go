// vybe-test: go/method_values/method_on_function_type
// origin: languages/go/tests/go/test_method_values.rs
// vybe-test-mode: compile

package main
type F func()
func (f F) call() { f() }
func main() { var fn F = func() {}
_ = fn.call }
