// vybe-test: go/flag_parsing_extended/flag_string_set_overrides_default
// origin: languages/go/tests/go/test_flag_parsing_extended.rs

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
_ = flag.Set("mode", "prod")
__check(fmt.Sprint(*mode), "prod") }
