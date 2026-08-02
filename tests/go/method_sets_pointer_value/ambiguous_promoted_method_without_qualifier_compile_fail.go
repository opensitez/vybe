// vybe-test: go/method_sets_pointer_value/ambiguous_promoted_method_without_qualifier_compile_fail
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile-fail

package main
type a struct{}
func (a) f() {}
type b struct{}
func (b) f() {}
type c struct { a
b }
func main() { var x c
x.f() }
