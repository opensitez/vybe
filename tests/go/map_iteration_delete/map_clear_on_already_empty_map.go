// vybe-test: go/map_iteration_delete/map_clear_on_already_empty_map
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[int]int{}
clear(values)
__check(fmt.Sprint(len(values)), "0") }
