// vybe-test: go/mime_multipart_extended/multipart_writer_form_data_content_type_has_boundary
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
import "strings"
func main() { w := multipart.NewWriter(bytes.NewBuffer(nil))
ct := w.FormDataContentType()
_ = strings.Contains(ct, "boundary=") }
