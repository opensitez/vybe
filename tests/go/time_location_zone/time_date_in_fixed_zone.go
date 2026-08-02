// vybe-test: go/time_location_zone/time_date_in_fixed_zone
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

func main() { loc := time.FixedZone("PST", -8*3600)
t := time.Date(2022, 5, 10, 8, 0, 0, 0, loc)
utc := t.UTC()
__check(fmt.Sprint(utc.Hour()), "16") }
