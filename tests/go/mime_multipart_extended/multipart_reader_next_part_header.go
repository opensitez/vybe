// vybe-test: go/mime_multipart_extended/multipart_reader_next_part_header
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
w.CreateFormField("h")
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
p, _ := r.NextPart()
_ = p.Header }
