// vybe-test: go/cover_text_html_log/slog_log
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "context"
func main() { slog.Log(context.Background(), slog.LevelInfo, "msg") }
