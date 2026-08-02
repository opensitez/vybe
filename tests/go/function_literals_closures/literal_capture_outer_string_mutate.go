// vybe-test: go/function_literals_closures/literal_capture_outer_string_mutate
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := "a"
appendChar := func(c string) { s += c }
appendChar("b")
__check(fmt.Sprint(s), "ab") }
