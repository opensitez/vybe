// vybe-test: go/interfaces_patterns_extra/interface_pass_as_parameter_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type bird struct{}
func (bird) speak() string { return "chirp" }
func say(value speaker) string { return value.speak() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(say(bird{})), "chirp")
}
