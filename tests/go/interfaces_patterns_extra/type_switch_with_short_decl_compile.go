// vybe-test: go/interfaces_patterns_extra/type_switch_with_short_decl_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { switch value := interface{}(1)
value.(type) { case int: } }
