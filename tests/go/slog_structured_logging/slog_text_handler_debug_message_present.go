// vybe-test: go/slog_structured_logging/slog_text_handler_debug_message_present
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "bytes"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
h := slog.NewTextHandler(&buf, nil)
l := slog.New(h)
l.Debug("trace")
__check(fmt.Sprint(strings.Contains(buf.String(), "trace")), "true") }
