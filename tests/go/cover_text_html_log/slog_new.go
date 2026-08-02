// vybe-test: go/cover_text_html_log/slog_new
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "os"
func main() { _ = slog.New(slog.NewTextHandler(os.Stdout, nil)) }
