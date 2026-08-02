// vybe-test: go/map_iteration_delete/map_delete_one_key_preserves_remaining_lookups
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"keep": 10, "drop": 20, "stay": 30}
delete(values, "drop")
__check(fmt.Sprint(len(values)), "2")
__check(fmt.Sprint(values["keep"]), "10")
__check(fmt.Sprint(values["stay"]), "30") }
