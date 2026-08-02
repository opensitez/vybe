// vybe-test: go/cover_encoding_extra/ascii85_decoder_read
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
import "bytes"
func main() { r := ascii85.NewDecoder(bytes.NewBufferString("<~00~>"))
buf := make([]byte, 4)
_, _ = r.Read(buf) }
