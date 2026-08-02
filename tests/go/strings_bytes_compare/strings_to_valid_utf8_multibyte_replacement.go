// vybe-test: go/strings_bytes_compare/strings_to_valid_utf8_multibyte_replacement
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.ToValidUTF8(string([]byte{0xc0}), "REPL") }
