// vybe-test: go/errors_package/fmt_errorf_formats_without_wrap
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("code %d", 404)
__check(fmt.Sprint(err.Error()), "code 404") }
