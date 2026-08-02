// vybe-test: go/cover_encoding_extra/pem_get_line
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/pem"
func main() { line, rest := pem.GetLine([]byte("TYPE rest"))
_, _ = line, rest }
