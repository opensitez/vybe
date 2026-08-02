// vybe-test: go/cover_encoding_extra/xml_decoder_entity
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
func main() { d := xml.NewDecoder(strings.NewReader("<a/>"))
d.Entity = map[string]string{"amp": "&"}
_ = d.Entity["amp"] }
