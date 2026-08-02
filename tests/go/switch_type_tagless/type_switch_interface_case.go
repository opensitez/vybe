// vybe-test: go/switch_type_tagless/type_switch_interface_case
// origin: languages/go/tests/go/test_switch_type_tagless.rs
// vybe-test-mode: compile

package main
type P interface { M() }
func describe(v interface{}) { switch v.(type) { case P: _ = v } }
func main() {}
