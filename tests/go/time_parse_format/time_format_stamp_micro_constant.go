// vybe-test: go/time_parse_format/time_format_stamp_micro_constant
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

func main() { t := time.Date(2020, 3, 15, 14, 30, 5, 123456000, time.UTC)
__check(fmt.Sprint(t.Format(time.StampMicro)), "Mar 15 14:30:05.123456") }
