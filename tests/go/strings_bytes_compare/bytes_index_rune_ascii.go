// vybe-test: go/strings_bytes_compare/bytes_index_rune_ascii
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.IndexRune([]byte("go"), 'o') }
