// vybe-test: go/type_aliases/defined_types_same_underlying_mutual_cast
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type First int
type Second int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := Second(First(15))
__check(fmt.Sprint(value), "15") }
