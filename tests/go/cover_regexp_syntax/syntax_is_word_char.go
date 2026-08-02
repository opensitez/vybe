// vybe-test: go/cover_regexp_syntax/syntax_is_word_char
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { _ = syntax.IsWordChar(97) }
