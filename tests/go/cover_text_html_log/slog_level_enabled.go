// vybe-test: go/cover_text_html_log/slog_level_enabled
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "context"
func main() { _ = slog.Default().Enabled(context.Background(), slog.LevelInfo) }
