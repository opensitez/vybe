// vybe-test: go/encoding_xml_runtime/xml_encoder_decode_roundtrip_compile
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "bytes"
type T struct { X int `xml:"x"` }
func main() { buf := bytes.NewBuffer(nil)
e := xml.NewEncoder(buf)
e.Encode(T{X: 1})
d := xml.NewDecoder(buf)
var t T
_ = d.Decode(&t) }
