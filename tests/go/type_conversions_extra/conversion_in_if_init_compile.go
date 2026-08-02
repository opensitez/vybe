// vybe-test: go/type_conversions_extra/conversion_in_if_init_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { if value := int(9.8); value > 0 { _ = value } }
