// vybe-test: go/flag_parsing_extended/flag_duration_after_set_seconds
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

func main() { d := flag.Duration("timeout", 0, "")
_ = flag.Set("timeout", "2s")
__check(fmt.Sprint(*d), "2s") }
