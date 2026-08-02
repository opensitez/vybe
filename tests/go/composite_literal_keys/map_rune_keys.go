// vybe-test: go/composite_literal_keys/map_rune_keys
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[rune]string{'a': "alpha", 'z': "omega"}
__check(fmt.Sprint(m['a']), "alpha")
__check(fmt.Sprint(m['z']), "omega")
}
