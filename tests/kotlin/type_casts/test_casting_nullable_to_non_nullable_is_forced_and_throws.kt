// vybe-test: kotlin/type_casts/test_casting_nullable_to_non_nullable_is_forced_and_throws
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

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
            val source: Any? = null
            val direct: String? = source as String?
            __p((direct == null).toString())

            try {
                val strict: String = source as String
                __p((strict).toString())
            } catch (e: Exception) {
                __p(("caught").toString())
            }
        
__check("true\ncaught")
}
