// vybe-test: go/slices_delete_insert/slices_compact_func
// origin: languages/go/tests/go/test_slices_delete_insert.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []string{"a","","b"}
_ = slices.CompactFunc(s, func(a string) bool { return a == "" }) }
