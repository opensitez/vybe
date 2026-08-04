// vybe-test: go/time_parse_format/time_parse_duration_hours_minutes
// origin: languages/go/tests/go/test_time_parse_format.rs

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

func main() { d, _ := time.ParseDuration("2h30m")
__p(fmt.Sprint(d.Minutes())) 
__check("150")
}
