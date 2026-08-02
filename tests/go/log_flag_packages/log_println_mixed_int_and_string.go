// vybe-test: go/log_flag_packages/log_println_mixed_int_and_string
// origin: languages/go/tests/go/test_log_flag_packages.rs

package main
import "fmt"
import "log"
func main() { log.Println("n=", 42)
fmt.Println("tail") }
