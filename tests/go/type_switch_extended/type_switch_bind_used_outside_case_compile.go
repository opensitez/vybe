// vybe-test: go/type_switch_extended/type_switch_bind_used_outside_case_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
func tag(v interface{}) { switch t := v.(type) { case int: _ = t }
_ = t }
func main() { tag(1) }
