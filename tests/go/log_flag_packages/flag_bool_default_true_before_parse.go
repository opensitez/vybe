// vybe-test: go/log_flag_packages/flag_bool_default_true_before_parse
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

func main() { verbose := flag.Bool("verbose", true, "verbose mode")
__check(fmt.Sprint(*verbose), "true") }
