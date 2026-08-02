// vybe-test: go/cover_encoding_extra/xml_decoder_decode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
type T struct { X int `xml:"x"` }
func main() { d := xml.NewDecoder(strings.NewReader("<T x=\"3\"></T>"))
var t T
_ = d.Decode(&t) }
