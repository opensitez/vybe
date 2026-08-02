// vybe-test: go/cover_encoding_extra/xml_encoder_encode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "bytes"
type T struct { X int `xml:"x"` }
func main() { e := xml.NewEncoder(bytes.NewBuffer(nil))
_ = e.Encode(T{X: 2}) }
