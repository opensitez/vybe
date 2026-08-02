// vybe-test: go/type_conversions_extra/conversion_in_switch_init_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { switch value := int(10.3); value { case 10: } }
