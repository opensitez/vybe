// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_minmax
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
            val u = uintArrayOf(3u, 1u, 4u)
            val b = ubyteArrayOf(9u, 2u, 7u)
            var minU = u[0]
            var maxB = b[0].toInt()
            var i = 1
            while (i < u.size) {
                if (u[i] < minU) { minU = u[i] }
                i += 1
            }
            i = 0
            while (i < b.size) {
                val v = b[i].toInt()
                if (v > maxB) { maxB = v }
                i += 1
            }
            __p((minU.toString()).toString())
            __p((maxB.toString()).toString())
        
__check("1\n9")
}
