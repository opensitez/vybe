// vybe-test: go/defer_lifo_extended/defer_runs_on_os_exit_path_not_called
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("defer"), "run")
__check(fmt.Sprint("run"), "defer")
}
