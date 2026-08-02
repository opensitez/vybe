// vybe-test: go/mime_multipart_extended/multipart_writer_create_part_twice
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
import "net/textproto"
func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
h := make(textproto.MIMEHeader)
w.CreatePart(h)
w.CreatePart(h)
w.Close() }
