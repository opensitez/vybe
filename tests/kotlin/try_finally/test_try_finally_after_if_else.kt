// vybe-test: kotlin/try_finally/test_try_finally_after_if_else
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(v: Int): Int {
        return if (v > 0) {
            try { v } finally { __p(("pos").toString()) }
        } else {
            try { -v } finally { __p(("neg").toString()) }
        }
    }
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

fun main() { __p((run(1)).toString())
__p((run(-2)).toString()) 
__check("pos\n1\nneg\n2")
}
