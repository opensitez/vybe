// vybe-test: go/method_sets_pointer_value/address_of_temp_for_pointer_method_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type cell struct { n int }
func (c *cell) set(v int) { c.n = v }
func main() { c := cell{}
c.set(1) }
