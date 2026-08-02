// vybe-test: go/type_conversions_extra/conversion_in_assignment_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { var value int
value = int(8.9)
_ = value }
