// vybe-test: go/sort_package/sort_strings_are_sorted
// origin: languages/go/tests/go/test_sort_package.rs
// vybe-test-mode: compile

package main
import "sort"
func main() { _ = sort.StringsAreSorted([]string{"a","b"}) }
