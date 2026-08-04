// vybe-test: go/path_filepath_package/filepath_split_parent_and_leaf
// origin: languages/go/tests/go/test_path_filepath_package.rs

package main
import "fmt"
import "path/filepath"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { dir, file := filepath.Split("/opt/bin/go")
__p(fmt.Sprint(dir))
__p(fmt.Sprint(file)) 
__check("/opt/bin/\ngo")
}
