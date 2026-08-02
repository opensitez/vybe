// vybe-test: go/embed_unsafe_size/embed_bytes_var
// origin: languages/go/tests/go/test_embed_unsafe_size.rs
// vybe-test-mode: compile

package main
import _ "embed"
//go:embed data.bin
var b []byte
func main() { _ = b }
