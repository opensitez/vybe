// vybe-test: go/time_location_zone/time_unix_micro_now
// origin: languages/go/tests/go/test_time_location_zone.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _ = time.UnixMicro(time.Now().UnixMicro()) }
