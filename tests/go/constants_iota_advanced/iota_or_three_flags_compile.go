// vybe-test: go/constants_iota_advanced/iota_or_three_flags_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( F0 = 1 << iota; F1; F2; All = F0 | F1 | F2 )
func main() { _ = All }
