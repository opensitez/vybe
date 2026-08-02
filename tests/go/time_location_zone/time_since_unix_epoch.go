// vybe-test: go/time_location_zone/time_since_unix_epoch
// origin: languages/go/tests/go/test_time_location_zone.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _ = time.Since(time.Unix(0, 0)) }
