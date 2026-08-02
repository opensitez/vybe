// vybe-test: go/generics_types/generic_alias_with_type_param
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
type List[T any] []T
func main() { var xs List[int]
_ = xs }
