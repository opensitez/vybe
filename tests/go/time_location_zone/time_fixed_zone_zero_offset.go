// vybe-test: go/time_location_zone/time_fixed_zone_zero_offset
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

func main() { loc := time.FixedZone("UTC+0", 0)
t := time.Date(2020, 1, 1, 0, 0, 0, 0, loc)
__check(fmt.Sprint(t.UTC().Hour()), "0") }
