// vybe-test: go/slog_structured_logging/slog_text_handler_add_source_option
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
opts := &slog.HandlerOptions{AddSource: true}
h := slog.NewTextHandler(&buf, opts)
l := slog.New(h)
l.Info("src")
__check(fmt.Sprint(len(buf.String()) > 0), "true") }
