// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_anonymous_pointer
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type Inner struct { V int }
type Outer struct { *Inner }
func main() { var o Outer
_ = json.Unmarshal([]byte(`{"V":9}`), &o) }
