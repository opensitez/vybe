// vybe-test: go/flag_parsing_extended/flag_duration_var_package_scope
// origin: languages/go/tests/go/test_flag_parsing_extended.rs

package main
import "fmt"
import "flag"
import "time"
var delay = flag.Duration("delay", time.Millisecond, "")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(*delay), "1ms") }
