// vybe-test: go/cover_encoding_extra/xml_escape
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "bytes"
func main() { _ = xml.Escape(bytes.NewBuffer(nil), []byte("<tag>")) }
