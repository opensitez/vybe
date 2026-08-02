// vybe-test: go/constants_iota_advanced/iota_combined_add_and_shift_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( V = 1<<iota + iota; W )
func main() { _, _ = V, W }
