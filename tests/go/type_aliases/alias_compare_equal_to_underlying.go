// vybe-test: go/type_aliases/alias_compare_equal_to_underlying
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

func main() { count := Count(10)
__check(fmt.Sprint(count == 10), "true") }
