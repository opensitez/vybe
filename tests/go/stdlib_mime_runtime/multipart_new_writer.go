// vybe-test: go/stdlib_mime_runtime/multipart_new_writer
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { _ = multipart.NewWriter(bytes.NewBuffer(nil)) }
