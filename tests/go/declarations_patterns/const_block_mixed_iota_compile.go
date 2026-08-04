// vybe-test: go/declarations_patterns/const_block_mixed_iota_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
const ( first = iota; second; third )
func main() { _, _, _ = first, second, third }
