// vybe-test: kotlin/kotlin_number_conversion_apis/test_nan_and_infinity_parse
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

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
            val nan = 0.0 / 0.0
            val inf = 1.0 / 0.0
            val ninf = -1.0 / 0.0
            __p((nan.isNaN()).toString())
            __p((inf.isInfinite()).toString())
            __p((ninf.isInfinite()).toString())
            __p(((inf + inf).isInfinite()).toString())
        
__check("true\ntrue\ntrue\ntrue")
}
