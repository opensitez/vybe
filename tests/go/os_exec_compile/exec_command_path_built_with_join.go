// vybe-test: go/os_exec_compile/exec_command_path_built_with_join
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os/exec"
import "path/filepath"
func main() { bin := filepath.Join("usr", "bin", "env")
_ = exec.Command(bin, "sh") }
