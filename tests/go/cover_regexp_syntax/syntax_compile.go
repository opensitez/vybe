// vybe-test: go/cover_regexp_syntax/syntax_compile
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { re, _ := syntax.Parse("a", 0)
_, _ = syntax.Compile(re) }
