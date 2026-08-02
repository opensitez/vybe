// vybe-test: go/slice_aliasing_extra/append_into_struct_field_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
type holder struct { values []int }
func main() { value := holder{}
value.values = append(value.values, 1)
_ = value }
