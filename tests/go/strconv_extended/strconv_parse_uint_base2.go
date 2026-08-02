// vybe-test: go/strconv_extended/strconv_parse_uint_base2
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseUint("1111", 2, 64) }
