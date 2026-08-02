// vybe-test: go/os_process_environ/os_getenv_missing_returns_empty
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

func main() { __check(fmt.Sprint(os.Getenv("VYBE_NONEXISTENT_VAR_XYZ") == ""), "true") }
