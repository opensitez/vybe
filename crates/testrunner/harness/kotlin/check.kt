// Vybe test harness — Kotlin.
//
// Real Kotlin source alongside harness/go/check.go and harness/js/check.js,
// the way test262's assert.js is JavaScript.
//
// A test's verdict is its EXIT CODE. `__check` prints its diagnostic BEFORE
// throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
// says nothing at all.
//
// Output is COLLECTED, not paired. The emitter rewrites every `println(x)`
// into `__p(x)` and compares the whole output once at the end of `main`.
// Pairing the i-th print with the i-th expected line cannot assert anything
// about a loop, and loops alone were 517 of Kotlin's 5,619 cases — the single
// largest hole in the suite.
//
// `__buf` is a String built by concatenation, NOT a StringBuilder. Calling a
// method on a bare top-level field receiver fails under Vybe with "undefined
// is not callable"; concatenation onto one has no such problem. The Java
// harness carries the same note for the same measured reason.
//
// `__p` takes a String that the CALL SITE already rendered: the emitter writes
// `__p((x).toString())`, not `__p(x)`. Rendering inside the harness is what the
// two obvious spellings get wrong under Vybe:
//
// * `__buf + o` concatenates onto a String VARIABLE, which renders a Boolean
//   as 1/0 (measured: `var s = ""; s = s + true` gives "1" while
//   `println("" + true)` gives "true").
// * `o.toString()` inside `__p` calls a method on a parameter declared `Any?`,
//   and a method resolves from the receiver's DECLARED type — the same reason
//   the Java harness cannot call one on a static field.
//
// At the call site the expression keeps its own static type, and `(x).toString()`
// is exactly what the previous per-print emitter compared against, so every
// rendering the suite already agreed with cargo on is preserved.

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}
