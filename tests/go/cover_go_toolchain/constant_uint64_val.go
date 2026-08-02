// vybe-test: go/cover_go_toolchain/constant_uint64_val
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/constant"
func main() { _, _ = constant.Uint64Val(constant.MakeUint64(1)) }
