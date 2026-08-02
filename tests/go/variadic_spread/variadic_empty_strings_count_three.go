// vybe-test: go/variadic_spread/variadic_empty_strings_count_three
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func count(words ...string) int { return len(words) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(count("", "", "")), "3")
}
