// vybe-test: go/time_location_zone/time_date_utc_components
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

func main() { t := time.Date(2024, time.March, 15, 14, 30, 45, 0, time.UTC)
__check(fmt.Sprint(t.Year()), "2024")
__check(fmt.Sprint(t.Month()), "March")
__check(fmt.Sprint(t.Day()), "15")
__check(fmt.Sprint(t.Hour()), "14") }
