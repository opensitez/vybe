// vybe-test: go/cover_encoding_extra/base32_new_encoder
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/base32"
import "bytes"
func main() { _ = base32.NewEncoder(base32.StdEncoding, bytes.NewBuffer(nil)) }
