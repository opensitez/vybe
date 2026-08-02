// vybe-test: go/json_unmarshal_advanced/json_unmarshal_unicode_in_object_key
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

func main() { var m map[string]int
json.Unmarshal([]byte(`{"\u006b":1}`), &m)
__check(fmt.Sprint(m["k"]), "1") }
