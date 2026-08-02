// vybe-test: go/nil_zero_semantics_extra/zero_value_struct_with_slice_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
type holder struct { values []int }
func main() { var h holder
_ = h }
