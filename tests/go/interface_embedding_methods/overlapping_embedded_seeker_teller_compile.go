// vybe-test: go/interface_embedding_methods/overlapping_embedded_seeker_teller_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type seeker interface { tell() int }
type pointer interface { tell() int }
type locator interface { seeker
pointer }
type cursor struct { pos int }
func (c cursor) tell() int { return c.pos }
func main() { var loc locator = cursor{pos: 0}
_ = loc.tell() }
