// vybe-test: go/constants_iota_advanced/iota_three_const_blocks_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( A = iota )
const ( B = iota )
const ( C = iota )
func main() { _, _, _ = A, B, C }
