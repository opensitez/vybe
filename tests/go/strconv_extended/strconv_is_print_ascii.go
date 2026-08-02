// vybe-test: go/strconv_extended/strconv_is_print_ascii
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.IsPrint('A') }
