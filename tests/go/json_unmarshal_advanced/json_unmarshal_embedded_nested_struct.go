// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_nested_struct
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Addr struct { City string }
type Person struct { Name string
Addr }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p Person
json.Unmarshal([]byte(`{"Name":"Ann","City":"Paris"}`), &p)
__check(fmt.Sprint(p.Name), "Ann")
__check(fmt.Sprint(p.Addr.City), "Paris") }
