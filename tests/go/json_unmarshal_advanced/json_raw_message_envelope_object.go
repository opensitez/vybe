// vybe-test: go/json_unmarshal_advanced/json_raw_message_envelope_object
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Envelope struct { Payload json.RawMessage }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var e Envelope
json.Unmarshal([]byte(`{"Payload":{"x":1}}`), &e)
var m map[string]int
json.Unmarshal(e.Payload, &m)
__check(fmt.Sprint(m["x"]), "1") }
