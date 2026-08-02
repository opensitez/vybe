// vybe-test: go/strconv_extended/strconv_atoi_whitespace_trim
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.Atoi("  42  ") }
