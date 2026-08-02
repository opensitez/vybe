// vybe-test: go/os_process_environ/os_lstat_current_dir_succeeds
// origin: languages/go/tests/go/test_os_process_environ.rs

package main
import "fmt"
import "os"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { _, err := os.Lstat(".")
__check(fmt.Sprint(err == nil), "true") }
