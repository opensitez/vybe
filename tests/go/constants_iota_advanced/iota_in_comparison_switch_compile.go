// vybe-test: go/constants_iota_advanced/iota_in_comparison_switch_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( A = iota; B; C )
func main() { switch B { case 1: _ = A + C } }
