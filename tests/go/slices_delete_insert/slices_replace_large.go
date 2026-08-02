// vybe-test: go/slices_delete_insert/slices_replace_large
// origin: languages/go/tests/go/test_slices_delete_insert.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []int{1,2,3,4,5}
_ = slices.Replace(s, 1, 4, 9, 8) }
