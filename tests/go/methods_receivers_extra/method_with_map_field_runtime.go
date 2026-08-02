// vybe-test: go/methods_receivers_extra/method_with_map_field_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type bag struct { values map[string]int }
func (b bag) get(key string) int { return b.values[key] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{values: map[string]int{"a": 6}}
__check(fmt.Sprint(value.get("a")), "6")
}
