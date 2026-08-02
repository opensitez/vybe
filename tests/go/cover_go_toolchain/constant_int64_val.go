// vybe-test: go/cover_go_toolchain/constant_int64_val
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/constant"
func main() { _, _ = constant.Int64Val(constant.MakeInt64(1)) }
