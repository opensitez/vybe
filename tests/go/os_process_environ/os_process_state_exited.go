// vybe-test: go/os_process_environ/os_process_state_exited
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
import "os/exec"
func main() { cmd := exec.Command("true")
err := cmd.Run()
if err == nil { _ = cmd.ProcessState.Exited() } }
