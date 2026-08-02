// vybe-test: go/slice_copy_clear/clear_map_removes_entries
// origin: languages/go/tests/go/test_slice_copy_clear.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{"a":1,"b":2}
clear(m)
__check(fmt.Sprint(len(m)), "0") }
