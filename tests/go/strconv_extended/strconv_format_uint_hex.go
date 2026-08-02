// vybe-test: go/strconv_extended/strconv_format_uint_hex
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.FormatUint(255, 16) }
