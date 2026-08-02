// vybe-test: go/strings_unicode_extra/string_append_via_plus_equal_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { text := "go"
text += "lang"
__check(fmt.Sprint(text), "golang")
}
