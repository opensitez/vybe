// vybe-test: go/function_literals_closures/closure_type_in_slice_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { fns := []func(){ func() {}, func() {} }
_ = fns }
