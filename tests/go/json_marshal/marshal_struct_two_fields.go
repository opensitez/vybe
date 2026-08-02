// vybe-test: go/json_marshal/marshal_struct_two_fields
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

func main() { b, _ := json.Marshal(Person{Name: "Bob", Age: 30})
__check(fmt.Sprint(string(b)), "{\"Name\":\"Bob\",\"Age\":30}") }
