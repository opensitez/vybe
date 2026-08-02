// vybe-test: go/stdlib_encoding_misc/pem_encode
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/pem"
func main() { _ = pem.EncodeToMemory(&pem.Block{Type: "TEST", Bytes: []byte("x")}) }
