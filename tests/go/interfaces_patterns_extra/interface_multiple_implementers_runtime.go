// vybe-test: go/interfaces_patterns_extra/interface_multiple_implementers_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type dog struct{}
type cat struct{}
func (dog) speak() string { return "woof" }
func (cat) speak() string { return "meow" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []speaker{dog{}, cat{}}
__check(fmt.Sprint(values[0].speak()), "woof")
__check(fmt.Sprint(values[1].speak()), "meow")
}
