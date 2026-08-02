// vybe-test: go/function_literals_closures/closure_capture_array
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { arr := [2]int{1, 2}
read := func(i int) int { return arr[i] }
__check(fmt.Sprint(read(1)), "2") }
