// vybe-test: go/go_ast_parser_compile/scanner_init
// origin: languages/go/tests/go/test_go_ast_parser_compile.rs
// vybe-test-mode: compile

package main
import "go/scanner"
import "go/token"
func main() { var s scanner.Scanner
fs := token.NewFileSet()
f := fs.AddFile("x.go", fs.Base(), 10)
s.Init(f, []byte("package main"), nil, 0) }
