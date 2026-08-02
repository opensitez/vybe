// vybe-test: go/stdlib_encoding_misc/base32_encode
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/base32"
func main() { _ = base32.StdEncoding.EncodeToString([]byte("go")) }
