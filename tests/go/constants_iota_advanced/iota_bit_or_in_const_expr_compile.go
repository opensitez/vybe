// vybe-test: go/constants_iota_advanced/iota_bit_or_in_const_expr_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( A = 1 << iota; B; Mask = A | B )
func main() { _ = Mask }
