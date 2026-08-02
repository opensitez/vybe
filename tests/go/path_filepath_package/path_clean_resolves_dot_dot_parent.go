// vybe-test: go/path_filepath_package/path_clean_resolves_dot_dot_parent
// origin: languages/go/tests/go/test_path_filepath_package.rs

package main
import "fmt"
import "path"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(path.Clean("/a/b/../c")), "/a/c") }
