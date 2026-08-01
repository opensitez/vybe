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

package main

import "fmt"

// __check ends the program unless got equals want.
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}
