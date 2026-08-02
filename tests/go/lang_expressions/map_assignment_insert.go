// vybe-test: go/lang_expressions/map_assignment_insert
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{}
m["k"] = 4
__check(fmt.Sprint(m["k"]), "4") }
