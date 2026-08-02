// vybe-test: go/constants_iota_advanced/iota_offset_plus_iota_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( Base = 100 + iota; Next )
func main() { _, _ = Base, Next }
