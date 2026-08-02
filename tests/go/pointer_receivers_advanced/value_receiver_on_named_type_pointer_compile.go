// vybe-test: go/pointer_receivers_advanced/value_receiver_on_named_type_pointer_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type degree int
func (d degree) sign() int { if d < 0 { return -1 }
if d > 0 { return 1 }
return 0 }
func main() { var value *degree
_ = value.sign() }
