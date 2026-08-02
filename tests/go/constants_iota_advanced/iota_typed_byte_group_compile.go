// vybe-test: go/constants_iota_advanced/iota_typed_byte_group_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( B0 byte = iota; B1 )
func main() { _, _ = B0, B1 }
