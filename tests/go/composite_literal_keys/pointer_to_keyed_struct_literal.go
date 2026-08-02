// vybe-test: go/composite_literal_keys/pointer_to_keyed_struct_literal
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type metric struct { name string
value int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := &metric{value: 42, name: "latency"}
__check(fmt.Sprint(m.name), "latency")
__check(fmt.Sprint(m.value), "42")
}
