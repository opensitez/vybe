// vybe-test: go/slog_structured_logging/slog_text_handler_level_filter_error
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
opts := &slog.HandlerOptions{Level: slog.LevelError}
h := slog.NewTextHandler(&buf, opts)
l := slog.New(h)
l.Info("hidden")
__check(fmt.Sprint(buf.Len() == 0), "true") }
