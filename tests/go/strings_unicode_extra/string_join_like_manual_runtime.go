// vybe-test: go/strings_unicode_extra/string_join_like_manual_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { left, right := "vy", "be"
__check(fmt.Sprint(left + "-" + right), "vy-be")
}
