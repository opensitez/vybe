// vybe-test: go/slog_structured_logging/slog_text_handler_replace_attr
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "bytes"
func main() { opts := &slog.HandlerOptions{ReplaceAttr: func(groups []string, a slog.Attr) slog.Attr { return a }}
_ = slog.NewTextHandler(bytes.NewBuffer(nil), opts) }
