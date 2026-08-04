// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

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
            val u = uintArrayOf(7u, 8u, 9u)
            val b = ubyteArrayOf(1u, 2u, 3u)
            var found = false
            var missing = false
            for (x in u) { if (x == 8u) found = true }
            for (x in b) { if (x == 9u) missing = true }
            __p((found.toString()).toString())
            __p((missing.toString()).toString())
        
__check("true\nfalse")
}
