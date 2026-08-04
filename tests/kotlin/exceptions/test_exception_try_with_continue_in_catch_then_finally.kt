// vybe-test: kotlin/exceptions/test_exception_try_with_continue_in_catch_then_finally
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

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

fun main() {
    for (value in 1..4) {
        try {
            if (value == 2) {
                throw Exception("bad")
            }
            __p((value).toString())
            continue
        } catch (e: Exception) {
            __p(("caught").toString())
            continue
        } finally {
            __p(("finally").toString())
        }
    }
    __p(("done").toString())

__check("1\nfinally\ncaught\nfinally\n3\nfinally\n4\nfinally\ndone")
}
