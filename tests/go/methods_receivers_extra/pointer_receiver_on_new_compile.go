// vybe-test: go/methods_receivers_extra/pointer_receiver_on_new_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c *counter) bump() { c.n++ }
func main() { value := new(counter)
value.bump() }
