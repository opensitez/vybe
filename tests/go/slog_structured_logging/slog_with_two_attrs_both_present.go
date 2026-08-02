// vybe-test: go/slog_structured_logging/slog_with_two_attrs_both_present
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
l := slog.New(slog.NewTextHandler(&buf, nil)).With("a", 1, "b", 2)
l.Info("pair")
s := buf.String()
__check(fmt.Sprint(strings.Contains(s, "1")), "true")
__check(fmt.Sprint(strings.Contains(s, "2")), "true") }
