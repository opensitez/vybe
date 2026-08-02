// vybe-test: go/go_ast_parser_compile/ast_filter
// origin: languages/go/tests/go/test_go_ast_parser_compile.rs
// vybe-test-mode: compile

package main
import "go/ast"
import "go/parser"
func main() { f, _ := parser.ParseFile(nil, "", "package main", 0)
ast.FilterFile(f, func(s string) bool { return true }) }
