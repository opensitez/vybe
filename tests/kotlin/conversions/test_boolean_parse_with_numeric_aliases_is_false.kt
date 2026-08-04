// vybe-test: kotlin/conversions/test_boolean_parse_with_numeric_aliases_is_false
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

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
            __p(("1".toBoolean()).toString())
            __p(("0".toBoolean()).toString())
            __p(("TRUE ".toBoolean()).toString())
            __p((" false ".toBoolean()).toString())
        
__check("false\nfalse\nfalse\nfalse")
}
