// vybe-test: go/map_iteration_delete/map_clear_builtin_drains_three_entries
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"a": 1, "b": 2, "c": 3}
clear(values)
__check(fmt.Sprint(len(values)), "0") }
