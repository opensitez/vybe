// vybe-test: go/mime_multipart_extended/mime_type_by_extension_css
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _ = mime.TypeByExtension(".css") }
