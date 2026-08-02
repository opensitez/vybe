// vybe-test: go/type_switch_extended/type_switch_case_nil_type_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
func tag(v interface{}) { switch v.(type) { case nil: _ = v } }
func main() { tag(nil) }
