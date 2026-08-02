// vybe-test: go/encoding_xml_runtime/xml_decoder_raw_token_loop
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
func main() { d := xml.NewDecoder(strings.NewReader("<a><b/></a>"))
for { _, err := d.RawToken()
if err != nil { break } } }
