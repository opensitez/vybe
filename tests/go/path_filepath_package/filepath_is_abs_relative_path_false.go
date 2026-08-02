// vybe-test: go/path_filepath_package/filepath_is_abs_relative_path_false
// origin: languages/go/tests/go/test_path_filepath_package.rs

package main
import "fmt"
import "path/filepath"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(filepath.IsAbs("local")), "false") }
