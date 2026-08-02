// vybe-test: go/log_flag_packages/log_print_two_strings_no_newline
// origin: languages/go/tests/go/test_log_flag_packages.rs

package main
import "fmt"
import "log"
func main() { log.Print("a", "b")
fmt.Println("end") }
