// vybe-test: go/strings_bytes_compare/bytes_compare_both_nil
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.Compare(nil, nil) }
