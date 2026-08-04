// vybe-test: kotlin/exceptions/test_try_expression_finally_with_no_catch
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun status(flag: Boolean): String {
            return try {
                if (flag) "ok" else throw Exception("bad")
            } finally {
                __p(("cleanup").toString())
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

fun main() {
            try {
                __p((status(false)).toString())
            } catch (e: Exception) {
                __p((e.message).toString())
            }
        
__check("cleanup\nbad")
}
