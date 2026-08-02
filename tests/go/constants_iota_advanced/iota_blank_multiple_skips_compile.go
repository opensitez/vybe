// vybe-test: go/constants_iota_advanced/iota_blank_multiple_skips_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( _ = iota; _; _; V )
func main() { _ = V }
