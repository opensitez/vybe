// vybe-test: go/strings_bytes_compare/bytes_has_prefix_nil_safety
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.HasPrefix(nil, []byte{}) }
