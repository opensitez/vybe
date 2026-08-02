// vybe-test: go/embed_unsafe_size/embed_string_var
// origin: languages/go/tests/go/test_embed_unsafe_size.rs
// vybe-test-mode: compile

package main
import _ "embed"
//go:embed hello
var s string
func main() { _ = s }
