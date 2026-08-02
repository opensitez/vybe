// vybe-test: go/time_location_zone/time_fixed_zone_half_hour
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

func main() { loc := time.FixedZone("IST", 5*3600+30*60)
t := time.Date(2021, 7, 1, 10, 0, 0, 0, loc)
__check(fmt.Sprint(t.Hour()), "10") }
