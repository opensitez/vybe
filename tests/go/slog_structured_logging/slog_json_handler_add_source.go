// vybe-test: go/slog_structured_logging/slog_json_handler_add_source
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "bytes"
func main() { opts := &slog.HandlerOptions{AddSource: true}
_ = slog.NewJSONHandler(&bytes.Buffer{}, opts) }
