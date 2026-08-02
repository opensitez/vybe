// vybe-test: go/path_filepath_package/filepath_split_parent_and_leaf
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

func main() { dir, file := filepath.Split("/opt/bin/go")
__check(fmt.Sprint(dir), "/opt/bin/")
__check(fmt.Sprint(file), "go") }
