// vybe-test: go/cover_regexp_syntax/syntax_regexp_cap_names
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { re, _ := syntax.Parse("(a)", syntax.Perl)
_ = re.CapNames() }
