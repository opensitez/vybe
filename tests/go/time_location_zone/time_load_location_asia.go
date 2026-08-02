// vybe-test: go/time_location_zone/time_load_location_asia
// origin: languages/go/tests/go/test_time_location_zone.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _, _ = time.LoadLocation("Asia/Tokyo") }
