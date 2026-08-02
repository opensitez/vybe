// vybe-test: go/flag_parsing_extended/flag_int_set_negative
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

func main() { offset := flag.Int("offset", 0, "")
_ = flag.Set("offset", "-3")
__check(fmt.Sprint(*offset), "-3") }
