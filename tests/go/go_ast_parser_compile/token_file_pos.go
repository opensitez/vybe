// vybe-test: go/go_ast_parser_compile/token_file_pos
// origin: languages/go/tests/go/test_go_ast_parser_compile.rs
// vybe-test-mode: compile

package main
import "go/token"
func main() { fs := token.NewFileSet()
f := fs.AddFile("a.go", fs.Base(), 100)
_ = f.Pos(1) }
