// vybe-test: go/stdlib_mime_runtime/mime_type_by_extension
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "mime"
func main() { _ = mime.TypeByExtension(".html") }
