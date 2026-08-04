// vybe-test: kotlin/conversions/test_string_to_long_radix_negative_and_large
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
            __p(("7fffffff".toLong(16)).toString())
            __p(("-100000000".toLong(2)).toString())
            __p(("1fffffffffffff".toLong(16)).toString())
        
__check("2147483647\n-256\n9007199254740991")
}
