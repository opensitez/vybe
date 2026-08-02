// vybe-test: go/pointer_receivers_advanced/address_of_int_field_mutate_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type pair struct { x int
y int }
func main() { value := pair{x: 1, y: 2}
ptr := &value.x
*ptr = 9
_ = value }
