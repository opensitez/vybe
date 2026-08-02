// vybe-test: go/methods_receivers_extra/pointer_method_on_struct_literal_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c *counter) bump() { c.n++ }
func main() { value := &counter{}
value.bump() }
