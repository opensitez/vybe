// vybe-test: go/pointer_receivers_advanced/nil_pointer_value_receiver_field_deref_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type node struct { id int }
func (n node) read() int { return n.id }
func main() { var value *node
_ = value.read() }
