// vybe-test: go/strings_bytes_compare/bytes_equal_empty_slices
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.Equal([]byte{}, []byte{}) }
