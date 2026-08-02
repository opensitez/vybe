// vybe-test: go/pointer_receivers_advanced/address_of_field_in_pointer_receiver_method_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type pair struct { x int
y int }
func (p *pair) swap() { px := &p.x
py := &p.y
*px, *py = *py, *px }
func main() { value := pair{x: 1, y: 2}
value.swap() }
