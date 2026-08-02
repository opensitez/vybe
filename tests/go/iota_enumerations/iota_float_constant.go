// vybe-test: go/iota_enumerations/iota_float_constant
// origin: languages/go/tests/go/test_iota_enumerations.rs
// vybe-test-mode: compile

package main
const ( F = 1.0 + float64(iota); G )
func main() { _ = G }
