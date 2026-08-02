// vybe-test: go/go_ast_parser_compile/ast_walk
// origin: languages/go/tests/go/test_go_ast_parser_compile.rs
// vybe-test-mode: compile

package main
import "go/ast"
import "go/parser"
import "go/token"
func main() { fs := token.NewFileSet()
f, _ := parser.ParseFile(fs, "", "package main", 0)
ast.Walk(nil, f) }
