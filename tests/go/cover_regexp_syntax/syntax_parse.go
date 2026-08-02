// vybe-test: go/cover_regexp_syntax/syntax_parse
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { _, _ = syntax.Parse("a", syntax.Perl) }
