// vybe-test: go/blank_identifier_extended/blank_type_switch_discard_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { var v interface{} = "x"
switch v.(type) { case string: _ = v } }
