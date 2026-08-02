// vybe-test: go/strconv_package/strconv_format_complex
// origin: languages/go/tests/go/test_strconv_package.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.FormatComplex(1+2i, 'f', 2, 64) }
