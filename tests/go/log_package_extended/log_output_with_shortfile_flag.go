// vybe-test: go/log_package_extended/log_output_with_shortfile_flag
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetFlags(log.Lshortfile)
_ = log.Output(0, "o\n") }
