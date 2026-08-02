// vybe-test: go/constants_iota_advanced/iota_parenthesized_list_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( X, Y = iota, iota + 10 )
func main() { _, _ = X, Y }
