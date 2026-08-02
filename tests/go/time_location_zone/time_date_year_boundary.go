// vybe-test: go/time_location_zone/time_date_year_boundary
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

func main() { t := time.Date(1999, time.December, 31, 23, 59, 59, 0, time.UTC)
__check(fmt.Sprint(t.Year()), "1999")
__check(fmt.Sprint(t.Month()), "December") }
