// vybe-test: go/slog_structured_logging/slog_uint64_attr_value
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "bytes"
import "strings"
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

func main() { var buf bytes.Buffer
l := slog.New(slog.NewTextHandler(&buf, nil))
l.Info("u", slog.Uint64("n", 5))
__p(fmt.Sprint(strings.Contains(buf.String(), "5"))) 
__check("true")
}
