// vybe-test: go/strconv_extended/strconv_parse_bool_whitespace
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.ParseBool("  true  ") }
