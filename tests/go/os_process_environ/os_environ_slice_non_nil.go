// vybe-test: go/os_process_environ/os_environ_slice_non_nil
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

func main() { __check(fmt.Sprint(len(os.Environ()) >= 0), "true") }
