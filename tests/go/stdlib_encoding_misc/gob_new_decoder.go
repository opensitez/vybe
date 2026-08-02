// vybe-test: go/stdlib_encoding_misc/gob_new_decoder
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { _ = gob.NewDecoder(bytes.NewBuffer(nil)) }
