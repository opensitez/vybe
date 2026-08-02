// vybe-test: go/time_location_zone/time_fixed_zone_name
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

func main() { loc := time.FixedZone("Custom", 3600)
__check(fmt.Sprint(loc.String()), "Custom") }
