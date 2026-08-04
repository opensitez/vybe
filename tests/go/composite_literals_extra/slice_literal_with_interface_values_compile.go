// vybe-test: go/composite_literals_extra/slice_literal_with_interface_values_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []interface{}{1, "two", true}
_ = values }
