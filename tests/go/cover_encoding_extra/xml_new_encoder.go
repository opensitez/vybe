// vybe-test: go/cover_encoding_extra/xml_new_encoder
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "bytes"
func main() { _ = xml.NewEncoder(bytes.NewBuffer(nil)) }
