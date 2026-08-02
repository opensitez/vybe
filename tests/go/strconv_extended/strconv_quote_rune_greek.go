// vybe-test: go/strconv_extended/strconv_quote_rune_greek
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.QuoteRune('λ') }
