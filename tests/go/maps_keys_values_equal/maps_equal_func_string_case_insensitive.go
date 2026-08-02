// vybe-test: go/maps_keys_values_equal/maps_equal_func_string_case_insensitive
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := map[int]string{1: "Go"}
b := map[int]string{1: "go"}
eq := maps.EqualFunc(a, b, func(x, y string) bool { return strings.EqualFold(x, y) })
__check(fmt.Sprint(eq), "true") }
