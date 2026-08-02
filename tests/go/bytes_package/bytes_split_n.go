// vybe-test: go/bytes_package/bytes_split_n
// origin: languages/go/tests/go/test_bytes_package.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.SplitN([]byte("a,b,c"), []byte(","), 2) }
