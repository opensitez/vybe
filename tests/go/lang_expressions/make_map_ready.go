// vybe-test: go/lang_expressions/make_map_ready
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := make(map[string]int)
m["a"] = 1
__check(fmt.Sprint(m["a"]), "1") }
