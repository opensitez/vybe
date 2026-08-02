// vybe-test: go/method_sets_pointer_value/dual_embedded_same_method_name_requires_qualifier_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type a struct{}
func (a) f() {}
type b struct{}
func (b) f() {}
type c struct { a
b }
func main() { var x c
x.a.f()
x.b.f() }
