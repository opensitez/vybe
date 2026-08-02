// vybe-test: go/time_location_zone/time_zone_offset_fixed
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

func main() { loc := time.FixedZone("X", 7200)
t := time.Date(2020, 1, 1, 0, 0, 0, 0, loc)
_, off := t.Zone()
__check(fmt.Sprint(off), "7200") }
