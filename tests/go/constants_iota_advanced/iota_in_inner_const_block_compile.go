// vybe-test: go/constants_iota_advanced/iota_in_inner_const_block_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( outer = 1; inner = iota; tail )
func main() { _, _ = inner, tail }
