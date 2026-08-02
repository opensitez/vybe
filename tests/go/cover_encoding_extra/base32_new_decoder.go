// vybe-test: go/cover_encoding_extra/base32_new_decoder
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/base32"
import "strings"
func main() { _ = base32.NewDecoder(base32.StdEncoding, strings.NewReader("MZXW6Y==")) }
