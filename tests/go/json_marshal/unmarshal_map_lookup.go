// vybe-test: go/json_marshal/unmarshal_map_lookup
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m map[string]int
json.Unmarshal([]byte("{\"key\":7}"), &m)
__check(fmt.Sprint(m["key"]), "7") }
