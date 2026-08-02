// vybe-test: go/cover_go_toolchain/doc_to_text
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/doc"
import "bytes"
func main() { var b bytes.Buffer
doc.ToText(&b, "Title", []byte("body"), "", 0) }
