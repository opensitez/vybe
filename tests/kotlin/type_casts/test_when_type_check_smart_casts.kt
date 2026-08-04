// vybe-test: kotlin/type_casts/test_when_type_check_smart_casts
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun describe(value: Any): String {
            return when {
                value is String -> "string:" + value.length
                value is Int -> "int:" + (value + 1)
                value is Boolean -> "bool:" + (if (value) 1 else 0)
                else -> "other"
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
            __p((describe("kotlin")).toString())
            __p((describe(6)).toString())
            __p((describe(false)).toString())
            __p((describe(1.5)).toString())
        
__check("string:6\nint:7\nbool:0\nother")
}
