// vybe-test: go/strings_unicode_extra/string_in_slice_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []string{"a", "b", "c"}
__check(fmt.Sprint(values[2]), "c")
}
