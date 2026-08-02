// vybe-test: go/time_location_zone/time_unix_zero_utc_year
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

func main() { t := time.Unix(0, 0).UTC()
__check(fmt.Sprint(t.Year()), "1970")
__check(fmt.Sprint(t.Month()), "January") }
