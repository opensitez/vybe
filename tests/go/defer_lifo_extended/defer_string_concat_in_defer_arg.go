// vybe-test: go/defer_lifo_extended/defer_string_concat_in_defer_arg
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := "go"
defer __check(fmt.Sprint(s + "lang"), "golang")
s = "rust"
}
