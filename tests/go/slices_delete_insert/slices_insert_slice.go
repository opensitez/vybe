// vybe-test: go/slices_delete_insert/slices_insert_slice
// origin: languages/go/tests/go/test_slices_delete_insert.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []int{1,4}
_ = slices.Insert(s, 1, []int{2,3}...) }
