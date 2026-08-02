// vybe-test: go/stdlib_encoding_misc/ascii85_encode
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
import "bytes"
func main() { _ = ascii85.NewEncoder(bytes.NewBuffer(nil)) }
