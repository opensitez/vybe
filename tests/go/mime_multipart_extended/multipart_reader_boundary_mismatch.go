// vybe-test: go/mime_multipart_extended/multipart_reader_boundary_mismatch
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
w.Close()
r := multipart.NewReader(&buf, "wrongBoundary")
_, err := r.NextPart()
_ = err }
