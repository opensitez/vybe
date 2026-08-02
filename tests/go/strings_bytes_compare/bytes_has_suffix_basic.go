// vybe-test: go/strings_bytes_compare/bytes_has_suffix_basic
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.HasSuffix([]byte("abc"), []byte("c")) }
