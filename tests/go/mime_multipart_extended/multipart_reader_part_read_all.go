// vybe-test: go/mime_multipart_extended/multipart_reader_part_read_all
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
import "io"
func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
fw, _ := w.CreateFormField("k")
fw.Write([]byte("v"))
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
p, _ := r.NextPart()
_, _ = io.ReadAll(p) }
