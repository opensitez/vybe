// vybe-test: go/methods_receivers_extra/pointer_receiver_interface_assign_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type adder interface { add(int) }
type counter struct { n int }
func (c *counter) add(v int) { c.n += v }
func main() { var value adder = &counter{}
_ = value }
