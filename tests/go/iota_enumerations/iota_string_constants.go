// vybe-test: go/iota_enumerations/iota_string_constants
// origin: languages/go/tests/go/test_iota_enumerations.rs
// vybe-test-mode: compile

package main
const ( A = "a"; B = iota )
func main() { _ = B }
