// vybe-test: go/time_parse_format/time_add_twenty_four_hours
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

func main() { t := time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC)
later := t.Add(24 * time.Hour)
__check(fmt.Sprint(later.Day()), "2") }
