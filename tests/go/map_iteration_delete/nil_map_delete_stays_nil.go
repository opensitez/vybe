// vybe-test: go/map_iteration_delete/nil_map_delete_stays_nil
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
delete(values, "x")
__check(fmt.Sprint(values == nil), "true") }
