// vybe-test: go/iota_enumerations/iota_with_explicit_value_then_iota
// origin: languages/go/tests/go/test_iota_enumerations.rs
// vybe-test-mode: compile

package main
const ( Start = 5; Next = iota; After )
func main() { _, _ = Next, After }
