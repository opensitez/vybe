// vybe-test: go/method_sets_pointer_value/pointer_to_value_still_gets_value_method_set_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type tile struct { color string }
func (t tile) hue() string { return t.color }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := tile{color: "red"}
p := &t
__check(fmt.Sprint(p.hue()), "red") }
