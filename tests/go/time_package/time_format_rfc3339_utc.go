// vybe-test: go/time_package/time_format_rfc3339_utc
// origin: languages/go/tests/go/test_time_package.rs

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

func main() { t := time.Date(2020, 1, 2, 3, 4, 5, 0, time.UTC)
__p(fmt.Sprint(t.Format(time.RFC3339))) 
__check("2020-01-02T03:04:05Z")
}
