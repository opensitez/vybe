// vybe-test: go/type_switch_extended/type_switch_fallthrough_invalid_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
func tag(v interface{}) { switch v.(type) { case int: fallthrough
case string: _ = v } }
func main() { tag(1) }
