// vybe-test: go/constants_iota_advanced/iota_second_group_after_non_iota_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const X = 9
const ( A = iota; B )
func main() { _, _ = A, B }
