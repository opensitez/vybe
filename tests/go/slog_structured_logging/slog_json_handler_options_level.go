// vybe-test: go/slog_structured_logging/slog_json_handler_options_level
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "bytes"
func main() { var buf bytes.Buffer
opts := &slog.HandlerOptions{Level: slog.LevelWarn}
l := slog.New(slog.NewJSONHandler(&buf, opts))
l.Warn("warn") }
