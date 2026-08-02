// vybe-test: go/constants_iota_advanced/iota_arithmetic_times_iota_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( X = iota * 5; Y; Z )
func main() { _, _, _ = X, Y, Z }
