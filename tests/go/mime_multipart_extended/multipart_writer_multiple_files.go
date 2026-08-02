// vybe-test: go/mime_multipart_extended/multipart_writer_multiple_files
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
w.CreateFormFile("a", "1.txt")
w.CreateFormFile("b", "2.txt")
w.Close() }
