// vybe-test: go/map_iteration_delete/nil_map_read_missing_key_returns_zero
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var values map[string]int
__check(fmt.Sprint(values["absent"]), "0") }
