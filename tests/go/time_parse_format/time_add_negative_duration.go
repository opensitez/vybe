// vybe-test: go/time_parse_format/time_add_negative_duration
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

func main() { t := time.Date(2020, 1, 2, 12, 0, 0, 0, time.UTC)
earlier := t.Add(-6 * time.Hour)
__check(fmt.Sprint(earlier.Hour()), "6") }
