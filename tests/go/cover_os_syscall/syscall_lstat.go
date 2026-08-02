// vybe-test: go/cover_os_syscall/syscall_lstat
// origin: languages/go/tests/go/test_cover_os_syscall.rs
// vybe-test-mode: compile

package main
import "syscall"
type StatT = syscall.Stat_t
func main() { var st StatT
_, _ = syscall.Lstat(".", &st) }
