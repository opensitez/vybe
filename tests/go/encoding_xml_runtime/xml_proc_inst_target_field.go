// vybe-test: go/encoding_xml_runtime/xml_proc_inst_target_field
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
func main() { p := xml.ProcInst{Target: "xml"}
_ = p.Target }
