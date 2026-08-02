// vybe-test: go/method_sets_pointer_value/pointer_only_method_interface_needs_pointer_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type editor interface { edit() }
type doc struct{}
func (d *doc) edit() {}
func main() { var e editor = &doc{}
_ = e }
