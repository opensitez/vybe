// Vybe test harness — Go.
//
// This is the Go equivalent of test262's harness/assert.js: ordinary source in
// the language under test, passed alongside the test file rather than compiled
// into the runner. `vybex a.go b.go` links several sources, and so does
// `go run a.go b.go`, so the same pair runs on both.
//
// A test's verdict is its EXIT CODE. That is the one mechanism shared by every
// language we target and every runtime we might compare against — C, COBOL and
// Fortran have no exceptions, but all of them can exit non-zero.
//
// __check prints its own diagnostic BEFORE it fails. That is not decoration:
// testecma relies on the thrown exception's message, and 1,692 of its 2,158
// failures come back as `RuntimeError: [object]` — no expected, no actual,
// nothing. A printed line survives that gap on every runtime.
//
// Output is COLLECTED, not paired. The emitter rewrites every `fmt.Println(x)`
// into `__p(...)` and compares the whole output once at the end of `main`.
// Pairing the i-th print with the i-th expected line cannot assert anything
// about a loop, and it also forced three other retreats that collection simply
// removes:
//
//   - `defer` ran the output out of source order. The buffer records the order
//     things ACTUALLY happened, so a defer anywhere — including one calling a
//     printing helper — now needs no reasoning at all.
//   - `fmt.Printf` writes many values per call. It becomes `__pr(fmt.Sprintf(…))`,
//     which is the same formatting by the same code.
//   - `fmt.Print` writes no newline, so calls shared a line. That is `__pr`.
//
// A goroutine is still out of reach: its output order is genuinely not knowable
// from the source, and collecting it does not make it so.
//
// `fmt.Sprintln` is NOT used — it is not implemented in Vybe (measured:
// "undefined is not callable"). Println's spacing is reproduced by the emitter
// instead, which is what the previous per-print rewriting already did.

package main

import "fmt"

var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}
