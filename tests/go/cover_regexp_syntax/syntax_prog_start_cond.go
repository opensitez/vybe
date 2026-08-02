// vybe-test: go/cover_regexp_syntax/syntax_prog_start_cond
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { re, _ := syntax.Parse("^a", 0)
p, _ := syntax.Compile(re)
_ = p.StartCond() }
