// vybe-test: go/cover_regexp_syntax/syntax_inst_match_rune
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { var i syntax.Inst
_ = i.MatchRune(97) }
