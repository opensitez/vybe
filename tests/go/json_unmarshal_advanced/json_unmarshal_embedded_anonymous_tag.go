// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_anonymous_tag
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Meta struct { Ver int `json:"ver"` }
type Doc struct { Meta
Title string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var d Doc
json.Unmarshal([]byte(`{"ver":2,"Title":"x"}`), &d)
__check(fmt.Sprint(d.Ver), "2")
__check(fmt.Sprint(d.Title), "x") }
