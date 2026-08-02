// vybe-test: go/interfaces_patterns_extra/interface_field_in_struct_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type dog struct{}
func (dog) speak() string { return "woof" }
type holder struct { value speaker }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := holder{value: dog{}}
__check(fmt.Sprint(h.value.speak()), "woof")
}
