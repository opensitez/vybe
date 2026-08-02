// vybe-test: go/mime_multipart_extended/mime_extensions_by_type_unknown
// origin: languages/go/tests/go/test_mime_multipart_extended.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _, err := mime.ExtensionsByType("application/x-unknown-vybe")
_ = err }
