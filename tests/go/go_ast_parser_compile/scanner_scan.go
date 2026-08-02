// vybe-test: go/go_ast_parser_compile/scanner_scan
// origin: languages/go/tests/go/test_go_ast_parser_compile.rs
// vybe-test-mode: compile

package main
import "go/scanner"
import "go/token"
func main() { var s scanner.Scanner
fs := token.NewFileSet()
f := fs.AddFile("x.go", fs.Base(), 20)
s.Init(f, []byte("package main"), nil, scanner.ScanComments)
_, _, _ = s.Scan() }
