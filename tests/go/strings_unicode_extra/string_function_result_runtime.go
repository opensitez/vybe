// vybe-test: go/strings_unicode_extra/string_function_result_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func label(v int) string { if v > 0 { return "pos" }
return "zero" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(label(1)), "pos")
}
