// vybe-test: go/strconv_extended/strconv_format_float_bits_32
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.FormatFloat(1.0, 'f', 2, 32) }
