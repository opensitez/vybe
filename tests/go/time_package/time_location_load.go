// vybe-test: go/time_package/time_location_load
// origin: languages/go/tests/go/test_time_package.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _, _ = time.LoadLocation("UTC") }
