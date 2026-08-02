// vybe-test: go/cover_runtime_testing/metrics_description
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "runtime/metrics"
func main() { desc := metrics.All()
if len(desc) > 0 { _ = desc[0].Name } }
