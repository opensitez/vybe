// vybe-test: go/map_iteration_delete/map_two_value_range_int_keys_lookup_sum
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[int]string{1: "a", 2: "b"}
__check(fmt.Sprint(len(values[1])), "1")
__check(fmt.Sprint(len(values[2])), "1") }
