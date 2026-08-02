// vybe-test: go/cover_go_toolchain/format_node
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/format"
import "go/parser"
func main() { f, _ := parser.ParseFile(nil, "", "package main", 0)
_ = format.Node(nil, f) }
