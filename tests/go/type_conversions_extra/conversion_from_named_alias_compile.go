// vybe-test: go/type_conversions_extra/conversion_from_named_alias_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
type level int
func main() { var value level = 12
_ = int(value) }
