// vybe-test: go/cover_go_toolchain/printer_commented_node
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/printer"
import "go/ast"
func main() { _ = printer.CommentedNode{Node: &ast.Ident{Name: "x"}, Comments: nil} }
