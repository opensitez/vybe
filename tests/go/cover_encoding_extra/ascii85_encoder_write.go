// vybe-test: go/cover_encoding_extra/ascii85_encoder_write
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
import "bytes"
func main() { e := ascii85.NewEncoder(bytes.NewBuffer(nil))
_, _ = e.Write([]byte("go")) }
