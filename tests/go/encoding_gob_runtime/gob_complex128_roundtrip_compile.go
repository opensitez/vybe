// vybe-test: go/encoding_gob_runtime/gob_complex128_roundtrip_compile
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { var buf bytes.Buffer
_ = gob.NewEncoder(&buf).Encode(complex(1, 2))
var c complex128
_ = gob.NewDecoder(&buf).Decode(&c) }
