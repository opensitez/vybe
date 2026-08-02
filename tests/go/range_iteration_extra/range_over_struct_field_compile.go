// vybe-test: go/range_iteration_extra/range_over_struct_field_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
type holder struct { values []int }
func main() { value := holder{values: []int{1}}
for _, item := range value.values { _ = item } }
