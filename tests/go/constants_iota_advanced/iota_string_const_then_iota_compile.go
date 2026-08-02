// vybe-test: go/constants_iota_advanced/iota_string_const_then_iota_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( Label = "go"; Code = iota; Next )
func main() { _, _ = Code, Next }
