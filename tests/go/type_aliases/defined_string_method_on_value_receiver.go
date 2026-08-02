// vybe-test: go/type_aliases/defined_string_method_on_value_receiver
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Tag string
func (t Tag) Len() int { return len(string(t)) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Tag("go").Len()), "2") }
