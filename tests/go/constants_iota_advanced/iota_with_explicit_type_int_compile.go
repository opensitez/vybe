// vybe-test: go/constants_iota_advanced/iota_with_explicit_type_int_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( A int = iota; B int )
func main() { _, _ = A, B }
