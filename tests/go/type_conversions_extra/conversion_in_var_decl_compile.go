// vybe-test: go/type_conversions_extra/conversion_in_var_decl_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
var value = int(9.4)
func main() { _ = value }
