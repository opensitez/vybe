// vybe-test: go/pointer_receivers_advanced/address_of_nested_struct_field_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type outer struct { core inner }
func main() { value := outer{core: inner{count: 6}}
ptr := &value.core.count
*ptr = 11
_ = value }
