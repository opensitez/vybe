// vybe-test: go/cover_go_toolchain/printer_config_fprint
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/printer"
import "go/ast"
func main() { var c printer.Config
_ = c.Fprint(nil, nil, &ast.Ident{Name: "x"}) }
