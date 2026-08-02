// vybe-test: go/encoding_xml_runtime/xml_comment_type_compile
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
func main() { _ = xml.Comment("note") }
