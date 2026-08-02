// vybe-test: go/flag_parsing_extended/flag_bool_set_to_false
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

func main() { debug := flag.Bool("debug", true, "")
_ = flag.Set("debug", "false")
__check(fmt.Sprint(*debug), "false") }
