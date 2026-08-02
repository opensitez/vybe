// vybe-test: go/log_flag_packages/flag_multiple_defaults_independent
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

func main() { host := flag.String("host", "localhost", "")
port := flag.Int("port", 3000, "")
__check(fmt.Sprint(*host), "localhost")
__check(fmt.Sprint(*port), "3000") }
