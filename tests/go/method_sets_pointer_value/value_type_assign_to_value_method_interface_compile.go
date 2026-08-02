// vybe-test: go/method_sets_pointer_value/value_type_assign_to_value_method_interface_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type worker interface { work() int }
type task struct { id int }
func (t task) work() int { return t.id }
func main() { var w worker = task{id: 1}
_ = w }
