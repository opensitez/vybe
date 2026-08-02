// vybe-test: go/cover_debug/debug_buildinfo_module_version
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "runtime/debug"
func main() { info, _ := debug.ReadBuildInfo()
if info != nil { _ = info.Main.Version } }
