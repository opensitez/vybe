// vybe-test: go/time_parse_format/time_parse_unix_date_constant
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

func main() { t, _ := time.Parse(time.UnixDate, "Sun Mar 15 14:30:00 UTC 2020")
__check(fmt.Sprint(t.Day()), "15") }
