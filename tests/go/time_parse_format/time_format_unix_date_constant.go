// vybe-test: go/time_parse_format/time_format_unix_date_constant
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

func main() { t := time.Date(2020, 3, 15, 14, 30, 0, 0, time.UTC)
__p(fmt.Sprint(t.Format(time.UnixDate))) 
__check("Sun Mar 15 14:30:00 UTC 2020")
}
