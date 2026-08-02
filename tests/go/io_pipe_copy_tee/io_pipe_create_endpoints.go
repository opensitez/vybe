// vybe-test: go/io_pipe_copy_tee/io_pipe_create_endpoints
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
func main() { pr, pw := io.Pipe()
_ = pr
_ = pw }
