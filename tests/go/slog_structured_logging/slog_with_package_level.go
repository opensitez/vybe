// vybe-test: go/slog_structured_logging/slog_with_package_level
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { _ = slog.With("k", "v").With("k2", 2) }
