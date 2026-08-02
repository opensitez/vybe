// vybe-test: go/flag_parsing_extended/flag_set_int64_string
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

func main() { v := flag.Int64("v", 0, "")
_ = flag.Set("v", "9223372036854775807")
__check(fmt.Sprint(*v > 0), "true") }
