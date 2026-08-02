// vybe-test: go/cover_encoding_extra/base32_std_encoding_decode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/base32"
func main() { dst := make([]byte, 8)
_, _ = base32.StdEncoding.Decode(dst, []byte("MZXW6Y==")) }
