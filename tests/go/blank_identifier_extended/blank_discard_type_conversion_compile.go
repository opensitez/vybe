// vybe-test: go/blank_identifier_extended/blank_discard_type_conversion_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
type score int
func main() { var s score = 3
_ = int(s) }
