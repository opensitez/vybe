// vybe-test: go/cover_regexp_syntax/syntax_regexp_string
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { re, _ := syntax.Parse("a", 0)
_ = re.String() }
