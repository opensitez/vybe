// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_shadowed_field
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Base struct { ID int `json:"id"` }
type Ext struct { Base
ID int `json:"ext_id"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var e Ext
json.Unmarshal([]byte(`{"id":1,"ext_id":9}`), &e)
__check(fmt.Sprint(e.Base.ID), "1")
__check(fmt.Sprint(e.ID), "9") }
