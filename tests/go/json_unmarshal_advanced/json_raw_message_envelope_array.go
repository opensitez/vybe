// vybe-test: go/json_unmarshal_advanced/json_raw_message_envelope_array
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Envelope struct { Data json.RawMessage }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var e Envelope
json.Unmarshal([]byte(`{"Data":[1,2]}`), &e)
var s []int
json.Unmarshal(e.Data, &s)
__check(fmt.Sprint(len(s)), "2")
__check(fmt.Sprint(s[1]), "2") }
