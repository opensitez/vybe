// vybe-test: go/strconv_extended/strconv_parse_float_hex_p754
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseFloat("0x1.fp+2", 64) }
