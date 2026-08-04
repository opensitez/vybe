// vybe-test: kotlin/bitwise_operations/test_bit_masking_even_and_odd
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

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
            val mask = 1
            val values = listOf(1, 2, 3, 4, 5, 6, 7, 8)
            val onlyEven = values.filter { it and 1 == 0 }
            val onlyOdd = values.filter { it and 1 == 1 }
            __p((onlyEven.joinToString(",")).toString())
            __p((onlyOdd.joinToString(",")).toString())
        
__check("2,4,6,8\n1,3,5,7")
}
