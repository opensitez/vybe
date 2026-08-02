// vybe-test: go/iota_enumerations/iota_in_parenthesized_group
// origin: languages/go/tests/go/test_iota_enumerations.rs
// vybe-test-mode: compile

package main
const ( X, Y = iota, iota + 1 )
func main() { _, _ = X, Y }
