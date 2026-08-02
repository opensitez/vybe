// vybe-test: go/switch_fallthrough_extended/fallthrough_in_type_switch_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func tag(v interface{}) { switch v.(type) { case int: fallthrough
case string: _ = v } }
func main() { tag(1) }
