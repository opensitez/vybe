// vybe-test: go/strconv_extended/strconv_format_float_nan
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.FormatFloat(0.0/0.0, 'g', 0, 64) }
