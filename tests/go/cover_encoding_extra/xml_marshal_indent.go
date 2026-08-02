// vybe-test: go/cover_encoding_extra/xml_marshal_indent
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
type T struct { X int `xml:"x"` }
func main() { _, _ = xml.MarshalIndent(T{X: 1}, "", "  ") }
