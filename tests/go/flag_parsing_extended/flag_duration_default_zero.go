// vybe-test: go/flag_parsing_extended/flag_duration_default_zero
// origin: languages/go/tests/go/test_flag_parsing_extended.rs

package main
import "fmt"
import "flag"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { d := flag.Duration("timeout", 0, "")
__check(fmt.Sprint(*d == 0), "true") }
