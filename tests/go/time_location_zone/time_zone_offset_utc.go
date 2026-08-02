// vybe-test: go/time_location_zone/time_zone_offset_utc
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

func main() { t := time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC)
_, off := t.Zone()
__check(fmt.Sprint(off), "0") }
