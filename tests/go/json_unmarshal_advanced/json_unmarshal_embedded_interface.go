// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_interface
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type Base struct { N int }
type Ext struct { Base
Extra string }
func main() { var e Ext
_ = json.Unmarshal([]byte(`{"N":1,"Extra":"x"}`), &e) }
