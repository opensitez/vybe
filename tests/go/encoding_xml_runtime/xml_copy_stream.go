// vybe-test: go/encoding_xml_runtime/xml_copy_stream
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "bytes"
func main() { dst := bytes.NewBuffer(nil)
_ = xml.Copy(dst, bytes.NewBufferString("<root/>")) }
