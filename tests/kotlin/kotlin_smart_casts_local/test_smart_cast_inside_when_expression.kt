// vybe-test: kotlin/kotlin_smart_casts_local/test_smart_cast_inside_when_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_smart_casts_local.rs

fun score(value: Any): String = when (value) {
            is String -> "s:" + value.length
            is Double -> "d:" + value.toInt()
            is Boolean -> "b:" + value
            else -> "n"
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
            __p((score("abc")).toString())
            __p((score(4.9)).toString())
            __p((score(false)).toString())
        
__check("s:3\nd:4\nb:false")
}
