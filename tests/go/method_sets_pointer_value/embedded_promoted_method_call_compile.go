// vybe-test: go/method_sets_pointer_value/embedded_promoted_method_call_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) f() {}
type outer struct { inner }
func main() { var o outer
o.f() }
