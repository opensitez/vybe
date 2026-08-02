// vybe-test: go/cover_go_toolchain/doc_to_html
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/doc"
import "bytes"
func main() { var b bytes.Buffer
doc.ToHTML(&b, "Title", []byte("body")) }
