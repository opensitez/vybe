// vybe-test: go/constants_iota_advanced/iota_shift_with_iota_multiplier_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( S = 1 << (iota + 1); T )
func main() { _, _ = S, T }
