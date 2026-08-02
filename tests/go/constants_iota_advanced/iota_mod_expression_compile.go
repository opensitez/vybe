// vybe-test: go/constants_iota_advanced/iota_mod_expression_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( A = iota % 3; B; C; D )
func main() { _, _, _, _ = A, B, C, D }
