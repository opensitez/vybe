// vybe-test: go/variadic_advanced/variadic_pointer_receiver_method_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
type acc struct { n int }
func (a *acc) add(nums ...int) { for _, v := range nums { a.n += v } }
func main() { x := acc{}
x.add(1, 2)
_ = x.n }
