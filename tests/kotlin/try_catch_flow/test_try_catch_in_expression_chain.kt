// vybe-test: kotlin/try_catch_flow/test_try_catch_in_expression_chain
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun value(x: Int): Int {
            return try {
                if (x == 0) throw IllegalStateException()
                10 / x
            } catch (e: IllegalStateException) {
                -1
            } catch (e: Exception) {
                -2
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
            __p((value(0)).toString())
            __p((value(2)).toString())
        
__check("-1\n5")
}
