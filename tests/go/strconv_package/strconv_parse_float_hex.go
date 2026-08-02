// vybe-test: go/strconv_package/strconv_parse_float_hex
// origin: languages/go/tests/go/test_strconv_package.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseFloat("0x1.8p0", 64) }
