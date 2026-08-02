// vybe-test: go/slog_structured_logging/slog_handler_with_group_chain
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "bytes"
func main() { h := slog.NewTextHandler(bytes.NewBuffer(nil), nil)
_ = h.WithGroup("g") }
