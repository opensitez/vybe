// vybe-test: go/slice_copy_clear/full_slice_expression_max
// origin: languages/go/tests/go/test_slice_copy_clear.rs
// vybe-test-mode: compile

package main
func main() { a := []int{1,2,3}
_ = a[0:1:2] }
