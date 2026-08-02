// vybe-test: go/defer_lifo_extended/defer_after_short_var_in_if
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { if x := 1; x > 0 { defer __check(fmt.Sprint(x), "1")
}
}
