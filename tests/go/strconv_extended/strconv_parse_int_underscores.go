// vybe-test: go/strconv_extended/strconv_parse_int_underscores
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseInt("1_000", 10, 64) }
