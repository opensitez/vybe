// vybe-test: go/function_literals_closures/closure_capture_struct_field
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
type counter struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := counter{n: 0}
bump := func() { c.n++ }
bump()
bump()
__check(fmt.Sprint(c.n), "2") }
