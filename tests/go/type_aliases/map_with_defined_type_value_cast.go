// vybe-test: go/type_aliases/map_with_defined_type_value_cast
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Level int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]Level{"a": Level(20)}
__check(fmt.Sprint(int(values["a"])), "20") }
