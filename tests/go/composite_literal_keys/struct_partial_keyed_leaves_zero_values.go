// vybe-test: go/composite_literal_keys/struct_partial_keyed_leaves_zero_values
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type config struct { host string
port int
debug bool }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := config{port: 8080}
__check(fmt.Sprint(c.port), "8080")
__check(fmt.Sprint(c.debug), "false")
}
