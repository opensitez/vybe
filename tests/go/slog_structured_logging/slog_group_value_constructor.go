// vybe-test: go/slog_structured_logging/slog_group_value_constructor
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { _ = slog.Group("g", slog.Bool("b", true)) }
