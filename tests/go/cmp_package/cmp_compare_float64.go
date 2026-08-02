// vybe-test: go/cmp_package/cmp_compare_float64
// origin: languages/go/tests/go/test_cmp_package.rs
// vybe-test-mode: compile

package main
import "cmp"
func main() { _ = cmp.Compare(1.5, 2.0) }
