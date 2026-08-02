// vybe-test: go/constants_iota_advanced/iota_complex_expression_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( V = (iota + 2) * (iota + 1); W )
func main() { _, _ = V, W }
