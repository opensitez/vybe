// vybe-test: go/strings_bytes_compare/bytes_to_upper_empty
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.ToUpper([]byte{}) }
