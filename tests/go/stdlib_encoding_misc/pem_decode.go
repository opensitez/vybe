// vybe-test: go/stdlib_encoding_misc/pem_decode
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/pem"
func main() { _, _ = pem.Decode([]byte("-----BEGIN TEST-----\nx\n-----END TEST-----")) }
