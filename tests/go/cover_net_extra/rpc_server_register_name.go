// vybe-test: go/cover_net_extra/rpc_server_register_name
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Args struct { A, B int }
type Arith int
func (t *Arith) Mul(args *Args, reply *int) error { return nil }
func main() { s := rpc.NewServer()
s.RegisterName("Math", new(Arith)) }
