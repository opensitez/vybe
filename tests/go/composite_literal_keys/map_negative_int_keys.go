// vybe-test: go/composite_literal_keys/map_negative_int_keys
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[int]string{-1: "minus", 0: "zero", 1: "plus"}
__check(fmt.Sprint(m[-1]), "minus")
__check(fmt.Sprint(m[0]), "zero")
}
