// vybe-test: go/cover_go_toolchain/constant_float32_val
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/constant"
func main() { _, _ = constant.Float32Val(constant.MakeFloat64(1.0)) }
