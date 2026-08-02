// vybe-test: go/go_ast_parser_compile/parser_parse_expr
// origin: languages/go/tests/go/test_go_ast_parser_compile.rs
// vybe-test-mode: compile

package main
import "go/parser"
import "go/token"
func main() { _, _ = parser.ParseExpr("1+2") }
