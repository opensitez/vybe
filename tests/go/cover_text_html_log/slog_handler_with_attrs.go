// vybe-test: go/cover_text_html_log/slog_handler_with_attrs
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "os"
func main() { h := slog.NewTextHandler(os.Stdout, nil)
_ = h.WithAttrs(nil) }
