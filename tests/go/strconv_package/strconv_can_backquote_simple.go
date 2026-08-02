// vybe-test: go/strconv_package/strconv_can_backquote_simple
// origin: languages/go/tests/go/test_strconv_package.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.CanBackquote("abc") }
