// vybe-test: go/os_exec_compile/exec_command_name_from_filepath_base
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os/exec"
import "path/filepath"
func main() { _ = exec.Command(filepath.Base("/bin/echo"), "vybe") }
