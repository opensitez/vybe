// vybe-test: go/iter_package/iter_pull2_manual_iteration
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { seq := func(yield func(rune, rune) bool) { yield('a', 'b') }
next, stop := iter.Pull2(seq)
defer stop()
for { _, _, ok := next()
if !ok { break } } }
