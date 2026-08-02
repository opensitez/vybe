// vybe-test: go/type_switch_extended/type_switch_empty_body_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
func tag(v interface{}) { switch v.(type) { case int: } }
func main() { tag(1) }
