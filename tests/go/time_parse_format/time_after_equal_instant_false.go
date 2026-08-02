// vybe-test: go/time_parse_format/time_after_equal_instant_false
// origin: languages/go/tests/go/test_time_parse_format.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := time.Unix(42, 0)
__check(fmt.Sprint(t.After(t)), "false") }
