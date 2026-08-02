// vybe-test: go/encoding_gob_runtime/gob_channel_not_encodable_compile
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { ch := make(chan int)
_ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(ch) }
