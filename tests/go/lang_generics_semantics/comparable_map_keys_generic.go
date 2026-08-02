// vybe-test: go/lang_generics_semantics/comparable_map_keys_generic
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
func Keys[M ~map[K]V, K comparable, V any](m M) int { return len(m) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Keys(map[string]int{"a":1})), "1") }
