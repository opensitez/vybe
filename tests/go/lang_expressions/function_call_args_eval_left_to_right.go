// vybe-test: go/lang_expressions/function_call_args_eval_left_to_right
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func f(a, b int) int { return a*10+b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { i := 0
i++
__check(fmt.Sprint(f(i, i)), "12") }
