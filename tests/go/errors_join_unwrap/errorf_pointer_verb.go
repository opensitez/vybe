// vybe-test: go/errors_join_unwrap/errorf_pointer_verb
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 7
err := fmt.Errorf("ptr %p", &n)
__check(fmt.Sprint(len(err.Error()) > 0), "true") }
