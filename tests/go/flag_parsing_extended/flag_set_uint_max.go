// vybe-test: go/flag_parsing_extended/flag_set_uint_max
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

func main() { v := flag.Uint("v", 0, "")
_ = flag.Set("v", "4294967295")
__check(fmt.Sprint(*v), "4294967295") }
