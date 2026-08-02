// vybe-test: go/constants_iota_advanced/iota_rune_from_iota_offset_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( R rune = '0' + iota; S )
func main() { _, _ = R, S }
