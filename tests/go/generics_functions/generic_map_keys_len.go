// vybe-test: go/generics_functions/generic_map_keys_len
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
func KeysLen[K comparable, V any](m map[K]V) int { return len(m) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(KeysLen(map[string]int{"a":1})), "1") }
