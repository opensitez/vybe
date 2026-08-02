// vybe-test: go/type_aliases/alias_from_builtin_without_cast
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Count = int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { plain := 9
var count Count = plain
__check(fmt.Sprint(count), "9") }
