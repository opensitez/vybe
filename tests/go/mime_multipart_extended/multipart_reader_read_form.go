// vybe-test: go/mime_multipart_extended/multipart_reader_read_form
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
w.CreateFormField("a")
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
_, _ = r.ReadForm(1024) }
