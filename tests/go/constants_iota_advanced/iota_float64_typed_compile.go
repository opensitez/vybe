// vybe-test: go/constants_iota_advanced/iota_float64_typed_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( F float64 = iota; G )
func main() { _, _ = F, G }
