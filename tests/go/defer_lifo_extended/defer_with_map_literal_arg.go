// vybe-test: go/defer_lifo_extended/defer_with_map_literal_arg
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{"a": 1}
defer __check(fmt.Sprint(m["a"]), "1")
m["a"] = 9
}
