// vybe-test: go/flag_parsing_extended/flag_uint64_after_set
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

func main() { big := flag.Uint64("big", 0, "")
_ = flag.Set("big", "5000")
__check(fmt.Sprint(*big), "5000") }
