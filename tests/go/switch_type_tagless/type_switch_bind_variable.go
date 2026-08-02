// vybe-test: go/switch_type_tagless/type_switch_bind_variable
// origin: languages/go/tests/go/test_switch_type_tagless.rs
// vybe-test-mode: compile

package main
func describe(v interface{}) { switch value := v.(type) { case int: _ = value
default: _ = value } }
func main() { describe(1) }
