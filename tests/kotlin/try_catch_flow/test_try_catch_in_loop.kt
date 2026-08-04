// vybe-test: kotlin/try_catch_flow/test_try_catch_in_loop
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun maybe(i: Int): Int {
            if (i < 0) throw Exception("neg")
            return i
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
            var sum = 0
            for (i in -1..2) {
                try {
                    sum += maybe(i)
                } catch (e: Exception) {
                    sum += 10
                }
            }
            __p((sum).toString())
        
__check("12")
}
