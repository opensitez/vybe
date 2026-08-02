// vybe-test: go/type_aliases/alias_pointer_type_in_struct_field
// origin: languages/go/tests/go/test_type_aliases.rs
// vybe-test-mode: compile

package main
type IntPtr = *int
type holder struct { ptr IntPtr }
func main() { n := 1
_ = holder{ptr: &n} }
