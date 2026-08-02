// vybe-test: go/type_switch_extended/type_switch_duplicate_int_case_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
func tag(v interface{}) { switch v.(type) { case int: _ = v
case int: _ = v } }
func main() { tag(1) }
