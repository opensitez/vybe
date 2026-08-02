// vybe-test: go/os_process_environ/os_stat_current_dir_is_dir
// origin: languages/go/tests/go/test_os_process_environ.rs

package main
import "fmt"
import "os"
func main() { fi, err := os.Stat(".")
if err != nil { fmt.Println(false)
return }
fmt.Println(fi.IsDir()) }
