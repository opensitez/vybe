// vybe-test: go/cover_regexp_syntax/syntax_empty_op_context
// origin: languages/go/tests/go/test_cover_regexp_syntax.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { _ = syntax.EmptyOpContext(97, 98) }
