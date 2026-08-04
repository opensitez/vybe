// vybe-test: go/time_location_zone/time_date_in_fixed_zone
// origin: languages/go/tests/go/test_time_location_zone.rs

package main
import "fmt"
import "time"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { loc := time.FixedZone("PST", -8*3600)
t := time.Date(2022, 5, 10, 8, 0, 0, 0, loc)
utc := t.UTC()
__p(fmt.Sprint(utc.Hour())) 
__check("16")
}
