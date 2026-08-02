// vybe-test: go/cover_text_html_log/slog_log_attrs
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "context"
func main() { slog.LogAttrs(context.Background(), slog.LevelInfo, "msg", slog.Int("n", 1)) }
