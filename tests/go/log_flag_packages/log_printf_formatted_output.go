// vybe-test: go/log_flag_packages/log_printf_formatted_output
// origin: languages/go/tests/go/test_log_flag_packages.rs

package main
import "fmt"
import "log"
func main() { log.Printf("%s-%d", "vybe", 7)
fmt.Println("after") }
