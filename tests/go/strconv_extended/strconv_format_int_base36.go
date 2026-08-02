// vybe-test: go/strconv_extended/strconv_format_int_base36
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.FormatInt(35, 36) }
