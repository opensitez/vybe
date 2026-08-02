// vybe-test: go/method_sets_pointer_value/value_type_cannot_assign_pointer_only_interface_compile_fail
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile-fail

package main
type editor interface { edit() }
type doc struct{}
func (d *doc) edit() {}
func main() { var e editor = doc{}
_ = e }
