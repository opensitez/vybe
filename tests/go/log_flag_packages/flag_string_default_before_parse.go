// vybe-test: go/log_flag_packages/flag_string_default_before_parse
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

func main() { name := flag.String("name", "guest", "user name")
__check(fmt.Sprint(*name), "guest") }
