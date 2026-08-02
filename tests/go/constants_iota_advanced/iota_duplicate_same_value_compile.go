// vybe-test: go/constants_iota_advanced/iota_duplicate_same_value_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( P = iota; Q = P; R = iota )
func main() { _, _, _ = P, Q, R }
