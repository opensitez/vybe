// vybe-test: go/encoding_xml_runtime/xml_decoder_skip_nested
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
func main() { d := xml.NewDecoder(strings.NewReader("<a><b><c/></b></a>"))
d.Token()
_ = d.Skip() }
