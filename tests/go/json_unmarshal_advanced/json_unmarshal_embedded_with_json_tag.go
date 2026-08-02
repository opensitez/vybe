// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_with_json_tag
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type Meta struct { Tag string `json:"tag"` }
type Doc struct { Meta
Body string }
func main() { var d Doc
_ = json.Unmarshal([]byte(`{"tag":"v1","Body":"hi"}`), &d) }
