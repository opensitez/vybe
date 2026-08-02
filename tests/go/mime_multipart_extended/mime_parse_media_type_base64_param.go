// vybe-test: go/mime_multipart_extended/mime_parse_media_type_base64_param
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _, params, _ := mime.ParseMediaType("text/plain; charset=iso-8859-1")
_ = params["charset"] }
