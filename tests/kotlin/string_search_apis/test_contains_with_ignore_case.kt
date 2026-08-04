// vybe-test: kotlin/string_search_apis/test_contains_with_ignore_case
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

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
            val text = "Kotlin"
            __p((text.contains("kin")).toString())
            __p((text.contains("KIN", ignoreCase = true)).toString())
            __p((text.contains('K')).toString())
            __p((text.contains('z')).toString())
        
__check("true\ntrue\ntrue\nfalse")
}
