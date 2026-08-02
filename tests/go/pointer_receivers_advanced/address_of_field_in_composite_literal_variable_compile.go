// vybe-test: go/pointer_receivers_advanced/address_of_field_in_composite_literal_variable_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type row struct { cells [2]int }
func main() { value := row{cells: [2]int{1, 2}}
ptr := &value.cells[1]
_ = ptr }
