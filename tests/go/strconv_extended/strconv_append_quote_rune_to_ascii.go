// vybe-test: go/strconv_extended/strconv_append_quote_rune_to_ascii
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _, _ = strconv.AppendQuoteRuneToASCII([]byte{}, 'A') }
