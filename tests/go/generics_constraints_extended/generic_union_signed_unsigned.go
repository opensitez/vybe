// vybe-test: go/generics_constraints_extended/generic_union_signed_unsigned
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func PickNum[T int | uint](v T) T { return v }
func main() { _ = PickNum(-3)
_ = PickNum(uint(3)) }
