// vybe-test: go/interfaces_patterns_extra/interface_value_in_slice_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type dog struct{}
func (dog) speak() string { return "woof" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []speaker{dog{}}
__check(fmt.Sprint(values[0].speak()), "woof")
}
