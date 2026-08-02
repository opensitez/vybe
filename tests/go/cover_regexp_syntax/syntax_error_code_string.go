// vybe-test: go/cover_regexp_syntax/syntax_error_code_string
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { _ = syntax.ErrInvalidRepeatSize.String() }
