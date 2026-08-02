// vybe-test: go/strings_unicode_extra/string_concat_with_number_string_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { count := "3"
__check(fmt.Sprint("items:" + count), "items:3")
}
