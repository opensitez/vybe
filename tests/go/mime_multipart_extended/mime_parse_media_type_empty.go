// vybe-test: go/mime_multipart_extended/mime_parse_media_type_empty
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _, _, err := mime.ParseMediaType("")
_ = err }
