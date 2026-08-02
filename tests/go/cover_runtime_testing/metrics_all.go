// vybe-test: go/cover_runtime_testing/metrics_all
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/metrics"
func main() { _ = metrics.All() }
