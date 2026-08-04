// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_sum
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
            val u = uintArrayOf(1u, 2u, 3u)
            val b = ubyteArrayOf(10u, 20u)
            var uSum = 0u
            var bSum = 0u
            for (x in u) { uSum += x }
            for (x in b) { bSum += x.toUInt() }
            __p((uSum.toString()).toString())
            __p((bSum.toString()).toString())
        
__check("6\n30")
}
