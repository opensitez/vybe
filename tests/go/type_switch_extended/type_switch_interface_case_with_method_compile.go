// vybe-test: go/type_switch_extended/type_switch_interface_case_with_method_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
type reader interface { Read() int }
func tag(v interface{}) { switch v.(type) { case reader: _ = v } }
func main() { tag(1) }
