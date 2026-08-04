// vybe-test: kotlin/when_expressions/test_when_with_local_type_checks_and_smart_casts
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun convert(value: Any): String {
            return when (value) {
                is Int -> "i=" + value.toString()
                is Long -> "l=" + value.toString()
                is Double -> "d=" + value.toString()
                else -> "x"
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
            __p((convert(3)).toString())
            __p((convert(4L)).toString())
            __p((convert(1.5)).toString())
            __p((convert("x")).toString())
        
__check("i=3\nl=4\nd=1.5\nx")
}
