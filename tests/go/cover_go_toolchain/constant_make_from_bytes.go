// vybe-test: go/cover_go_toolchain/constant_make_from_bytes
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/constant"
func main() { _ = constant.MakeFromBytes([]byte{1}) }
