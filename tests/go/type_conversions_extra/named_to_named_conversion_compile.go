// vybe-test: go/type_conversions_extra/named_to_named_conversion_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type first int
type second int
func main() { var value first = 10
_ = second(value) }
