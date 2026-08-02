// vybe-test: go/constants_iota_advanced/iota_negative_expression_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( N = -1 - iota; M )
func main() { _, _ = N, M }
