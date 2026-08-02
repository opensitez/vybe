// vybe-test: go/constants_iota_advanced/iota_typed_rune_group_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( R0 rune = '!' + iota; R1 )
func main() { _, _ = R0, R1 }
