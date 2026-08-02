// vybe-test: go/bytes_package/bytes_replace_all
// origin: languages/go/tests/go/test_bytes_package.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.ReplaceAll([]byte("a.a"), []byte("."), []byte("-")) }
