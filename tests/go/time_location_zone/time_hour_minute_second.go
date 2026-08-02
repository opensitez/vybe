// vybe-test: go/time_location_zone/time_hour_minute_second
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

func main() { t := time.Date(2020, 6, 15, 9, 8, 7, 0, time.UTC)
__check(fmt.Sprint(t.Hour()), "9")
__check(fmt.Sprint(t.Minute()), "8")
__check(fmt.Sprint(t.Second()), "7") }
