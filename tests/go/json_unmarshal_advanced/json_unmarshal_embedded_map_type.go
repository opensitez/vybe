// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_map_type
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Meta map[string]string
type Doc struct { Meta
ID int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var d Doc
json.Unmarshal([]byte(`{"ID":1,"k":"v"}`), &d)
__check(fmt.Sprint(d.ID), "1")
__check(fmt.Sprint(d.Meta["k"]), "v") }
