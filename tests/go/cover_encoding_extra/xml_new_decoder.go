// vybe-test: go/cover_encoding_extra/xml_new_decoder
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "strings"
func main() { _ = xml.NewDecoder(strings.NewReader("<T/>")) }
