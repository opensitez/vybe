// vybe-test: go/strings_bytes_compare/strings_has_suffix_path_ext
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.HasSuffix("/tmp/x.tar.gz", ".gz") }
