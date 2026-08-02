// vybe-test: go/json_unmarshal_advanced/json_raw_message_marshal_passthrough
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { raw := json.RawMessage(`{"k":7}`)
b, _ := json.Marshal(raw)
__check(fmt.Sprint(string(b)), "{\"k\":7}") }
