// vybe-test: go/encoding_hex_base64/hex_dumper_writer
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "bytes"
import "encoding/hex"
func main() { var buf bytes.Buffer
w := hex.Dumper(&buf)
_, _ = w.Write([]byte("x"))
w.Close() }
