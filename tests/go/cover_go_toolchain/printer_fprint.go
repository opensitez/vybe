// vybe-test: go/cover_go_toolchain/printer_fprint
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/printer"
import "go/ast"
func main() { _ = printer.Fprint(nil, nil, &ast.Ident{Name: "x"}) }
