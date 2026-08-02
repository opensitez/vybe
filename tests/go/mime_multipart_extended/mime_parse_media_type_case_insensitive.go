// vybe-test: go/mime_multipart_extended/mime_parse_media_type_case_insensitive
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { mt, params, _ := mime.ParseMediaType("Text/HTML; Charset=UTF-8")
_ = mt
_ = params }
