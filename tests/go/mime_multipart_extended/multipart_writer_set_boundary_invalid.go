// vybe-test: go/mime_multipart_extended/multipart_writer_set_boundary_invalid
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime/multipart"
import "bytes"
func main() { w := multipart.NewWriter(bytes.NewBuffer(nil))
err := w.SetBoundary("bad boundary spaces")
_ = err }
