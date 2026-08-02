// vybe-test: go/flag_parsing_extended/flag_uint_after_set
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

func main() { count := flag.Uint("count", 0, "")
_ = flag.Set("count", "99")
__check(fmt.Sprint(*count), "99") }
