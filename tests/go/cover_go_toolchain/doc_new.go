// vybe-test: go/cover_go_toolchain/doc_new
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/doc"
import "go/parser"
import "go/token"
func main() { fs := token.NewFileSet()
f, _ := parser.ParseFile(fs, "", "package main\nfunc F() {}", 0)
_, _ = doc.New(f, "./", 0) }
