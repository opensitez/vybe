// vybe-test: go/constants_iota_advanced/iota_explicit_value_breaks_chain_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( Start = 5; Next = iota; After )
func main() { _, _ = Next, After }
