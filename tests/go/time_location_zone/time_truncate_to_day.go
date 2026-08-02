// vybe-test: go/time_location_zone/time_truncate_to_day
// origin: languages/go/tests/go/test_time_location_zone.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := time.Date(2023, 4, 5, 14, 35, 22, 0, time.UTC)
truncated := t.Truncate(24 * time.Hour)
__check(fmt.Sprint(truncated.Hour()), "0")
__check(fmt.Sprint(truncated.Day()), "5") }
