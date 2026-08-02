// vybe-test: go/pointer_receivers_advanced/new_pointer_passed_to_value_receiver_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type token struct { id int }
func (t token) value() int { return t.id }
func main() { created := new(token)
_ = created.value() }
