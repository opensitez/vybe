// vybe-test: go/composite_literal_keys/nested_struct_map_array_all_keyed
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type entry struct { scores []int }
type table struct { rows map[string]entry }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := table{rows: map[string]entry{"a": {scores: []int{0: 100, 2: 300}}}}
__check(fmt.Sprint(t.rows["a"].scores[0]), "100")
__check(fmt.Sprint(t.rows["a"].scores[2]), "300")
}
