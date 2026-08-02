// vybe-test: go/bytes_package/bytes_map
// origin: languages/go/tests/go/test_bytes_package.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { _ = bytes.Map(func(r rune) rune { return r }, []byte("abc")) }
