// vybe-test: go/methods_receivers_extra/value_receiver_interface_assign_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type totaler interface { total() int }
type counter struct { n int }
func (c counter) total() int { return c.n }
func main() { var value totaler = counter{}
_ = value }
