// vybe-test: go/time_parse_format/time_add_date_month_rollover
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

func main() { t := time.Date(2020, 1, 31, 0, 0, 0, 0, time.UTC)
later := t.AddDate(0, 1, 0)
__check(fmt.Sprint(int(later.Month())), "3") }
