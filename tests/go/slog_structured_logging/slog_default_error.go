// vybe-test: go/slog_structured_logging/slog_default_error
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { slog.Default().Error("e") }
