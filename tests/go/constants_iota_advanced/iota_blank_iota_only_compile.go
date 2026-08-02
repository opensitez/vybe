// vybe-test: go/constants_iota_advanced/iota_blank_iota_only_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( _ = iota; K = iota )
func main() { _ = K }
