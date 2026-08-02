// vybe-test: go/encoding_xml_runtime/xml_decoder_input_pos
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
func main() { d := xml.NewDecoder(strings.NewReader("<a/>"))
_, _ = d.InputPos() }
