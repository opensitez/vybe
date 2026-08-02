// vybe-test: go/stdlib_mime_runtime/mime_format_media_type
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _ = mime.FormatMediaType("text/html", map[string]string{"charset": "utf-8"}) }
