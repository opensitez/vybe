// vybe-test: go/cmp_package/cmp_less_float64
// origin: languages/go/tests/go/test_cmp_package.rs
// vybe-test-mode: compile

package main
import "cmp"
func main() { _ = cmp.Less(1.0, 2.0) }
