// vybe-test: go/cover_encoding_extra/xml_decoder_skip
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
func main() { d := xml.NewDecoder(strings.NewReader("<a><b/></a>"))
_ = d.Skip() }
