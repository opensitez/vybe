// vybe-test: go/bytes_package/bytes_repeat
// origin: languages/go/tests/go/test_bytes_package.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.Repeat([]byte("go"), 2) }
