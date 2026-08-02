// vybe-test: go/strconv_extended/strconv_parse_int_base36
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseInt("z", 36, 64) }
