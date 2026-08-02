// vybe-test: go/mime_multipart_extended/mime_format_media_type_multipart_mixed
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _ = mime.FormatMediaType("multipart/mixed", map[string]string{"boundary": "----WebKit"}) }
