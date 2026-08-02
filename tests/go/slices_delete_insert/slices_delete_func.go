// vybe-test: go/slices_delete_insert/slices_delete_func
// origin: languages/go/tests/go/test_slices_delete_insert.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []int{1,2,3}
_ = slices.DeleteFunc(s, func(v int) bool { return v == 2 }) }
