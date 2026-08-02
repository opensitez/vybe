// vybe-test: go/cmp_package/cmp_or_strings
// origin: languages/go/tests/go/test_cmp_package.rs
// vybe-test-mode: compile

package main
import "cmp"
func main() { _ = cmp.Or("", "ok") }
