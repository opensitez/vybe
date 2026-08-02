// vybe-test: go/map_iteration_delete/map_clear_builtin_then_reinsert_one
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"old": 9}
clear(values)
values["new"] = 4
__check(fmt.Sprint(len(values)), "1")
__check(fmt.Sprint(values["new"]), "4") }
