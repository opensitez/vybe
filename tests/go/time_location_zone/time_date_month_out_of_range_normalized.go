// vybe-test: go/time_location_zone/time_date_month_out_of_range_normalized
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

func main() { t := time.Date(2020, 13, 1, 0, 0, 0, 0, time.UTC)
__check(fmt.Sprint(t.Month()), "January")
__check(fmt.Sprint(t.Year()), "2021") }
