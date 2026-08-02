// vybe-test: go/flag_parsing_extended/flag_int_rebind_pointer_after_set
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

func main() { n := flag.Int("n", 1, "")
_ = flag.Set("n", "8")
*n = 9
__check(fmt.Sprint(*n), "9") }
