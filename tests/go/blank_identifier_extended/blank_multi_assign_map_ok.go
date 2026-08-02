// vybe-test: go/blank_identifier_extended/blank_multi_assign_map_ok
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{"k": 9}
v, _ := m["k"]
__check(fmt.Sprint(v), "9") }
