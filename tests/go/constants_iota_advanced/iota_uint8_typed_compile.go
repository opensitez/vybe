// vybe-test: go/constants_iota_advanced/iota_uint8_typed_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( U uint8 = iota; V )
func main() { _, _ = U, V }
