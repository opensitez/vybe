// vybe-test: go/time_parse_format/time_sub_same_instant_zero
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

func main() { a := time.Unix(1000, 0)
b := time.Unix(1000, 0)
__check(fmt.Sprint(a.Sub(b).Nanoseconds()), "0") }
