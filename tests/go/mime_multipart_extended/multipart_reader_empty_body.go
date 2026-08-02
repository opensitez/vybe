// vybe-test: go/mime_multipart_extended/multipart_reader_empty_body
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { r := multipart.NewReader(bytes.NewReader([]byte{}), "b")
_, err := r.NextPart()
_ = err }
