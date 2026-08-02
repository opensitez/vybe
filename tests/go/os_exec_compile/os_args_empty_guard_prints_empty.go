// vybe-test: go/os_exec_compile/os_args_empty_guard_prints_empty
// origin: languages/go/tests/go/test_os_exec_compile.rs

package main
import "fmt"
import "os"
func main() { if len(os.Args) > 0 { fmt.Println(os.Args[0]) } else { fmt.Println("empty") } }
