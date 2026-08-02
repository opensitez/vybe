// vybe-test: go/cover_regexp_syntax/syntax_error_string
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { var e syntax.Error
_ = e.Error() }
