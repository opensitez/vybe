// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_time_layout
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type When struct { Year int `json:"year"` }
type Event struct { When
Name string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var e Event
json.Unmarshal([]byte(`{"year":2024,"Name":"launch"}`), &e)
__check(fmt.Sprint(e.Year), "2024")
__check(fmt.Sprint(e.Name), "launch") }
