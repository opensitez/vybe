// vybe-test: go/log_package_extended/log_output_with_longfile_flag
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetFlags(log.Llongfile)
_ = log.Output(0, "o\n") }
