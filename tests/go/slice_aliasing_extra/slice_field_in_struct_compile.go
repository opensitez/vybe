// vybe-test: go/slice_aliasing_extra/slice_field_in_struct_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
type holder struct { values []int }
func main() { _ = holder{} }
