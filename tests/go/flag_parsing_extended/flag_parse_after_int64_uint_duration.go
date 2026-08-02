// vybe-test: go/flag_parsing_extended/flag_parse_after_int64_uint_duration
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.Int64("size", 0, "")
_ = flag.Uint("count", 0, "")
_ = flag.Duration("timeout", 0, "")
flag.Parse() }
