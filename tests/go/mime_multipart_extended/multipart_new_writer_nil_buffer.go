// vybe-test: go/mime_multipart_extended/multipart_new_writer_nil_buffer
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
func main() { _ = multipart.NewWriter(nil) }
