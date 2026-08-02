// vybe-test: go/time_location_zone/time_fixed_zone_positive_offset
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

func main() { loc := time.FixedZone("EST", -5*3600)
t := time.Date(2020, 1, 1, 12, 0, 0, 0, loc)
__check(fmt.Sprint(t.Location().String()), "EST")
__check(fmt.Sprint(t.Hour()), "12") }
