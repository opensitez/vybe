// vybe-test: go/json_marshal/unmarshal_struct_populates_fields
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Person struct { Name string
Age int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p Person
json.Unmarshal([]byte("{\"Name\":\"Bob\",\"Age\":30}"), &p)
__check(fmt.Sprint(p.Name), "Bob")
__check(fmt.Sprint(p.Age), "30") }
