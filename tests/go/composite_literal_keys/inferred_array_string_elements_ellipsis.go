// vybe-test: go/composite_literal_keys/inferred_array_string_elements_ellipsis
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { words := [...]string{"go", "vybe", "keys"}
__check(fmt.Sprint(len(words)), "3")
__check(fmt.Sprint(words[2]), "keys")
}
