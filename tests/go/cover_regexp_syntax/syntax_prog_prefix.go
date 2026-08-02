// vybe-test: go/cover_regexp_syntax/syntax_prog_prefix
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { re, _ := syntax.Parse("abc", 0)
p, _ := syntax.Compile(re)
_, _ = p.Prefix() }
