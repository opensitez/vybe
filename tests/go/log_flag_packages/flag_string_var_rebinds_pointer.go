// vybe-test: go/log_flag_packages/flag_string_var_rebinds_pointer
// origin: languages/go/tests/go/test_log_flag_packages.rs

package main
import "fmt"
import "flag"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { mode := flag.String("mode", "dev", "")
*mode = "prod"
__check(fmt.Sprint(*mode), "prod") }
