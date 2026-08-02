// vybe-test: go/os_exec_compile/os_args_len_is_zero
// origin: languages/go/tests/go/test_os_exec_compile.rs

package main
import "fmt"
import "os"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(len(os.Args)), "0") }
