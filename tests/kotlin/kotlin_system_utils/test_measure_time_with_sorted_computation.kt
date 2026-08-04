// vybe-test: kotlin/kotlin_system_utils/test_measure_time_with_sorted_computation
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

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
            val numbers = (1..20).toList().shuffled().sorted()
            val elapsed = kotlin.system.measureTimeMillis {
                __p((numbers.joinToString(",")).toString())
            }
            __p((numbers.size).toString())
            __p((elapsed >= 0).toString())
        
__check("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20\n20\ntrue")
}
