// vybe-test: go/slog_structured_logging/slog_json_handler_with_attrs
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "bytes"
func main() { var buf bytes.Buffer
l := slog.New(slog.NewJSONHandler(&buf, nil))
l.Info("json", slog.Int("n", 1)) }
