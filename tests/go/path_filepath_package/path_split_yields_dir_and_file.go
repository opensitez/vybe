// vybe-test: go/path_filepath_package/path_split_yields_dir_and_file
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

func main() { dir, file := path.Split("/a/b/c")
__check(fmt.Sprint(dir), "/a/b/")
__check(fmt.Sprint(file), "c") }
