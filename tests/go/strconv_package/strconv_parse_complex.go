// vybe-test: go/strconv_package/strconv_parse_complex
// origin: languages/go/tests/go/test_strconv_package.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseComplex("1+2i", 64) }
