// vybe-test: go/slog_structured_logging/slog_text_handler_level_filter_debug
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "bytes"
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
opts := &slog.HandlerOptions{Level: slog.LevelDebug}
h := slog.NewTextHandler(&buf, opts)
l := slog.New(h)
l.Debug("low")
__p(fmt.Sprint(buf.Len() > 0)) 
__check("true")
}
