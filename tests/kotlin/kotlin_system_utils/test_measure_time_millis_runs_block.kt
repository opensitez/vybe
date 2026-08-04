// vybe-test: kotlin/kotlin_system_utils/test_measure_time_millis_runs_block
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
            var seen = false
            val millis = kotlin.system.measureTimeMillis {
                seen = true
                var sum = 0
                for (i in 1..1000) sum += i
                __p((sum).toString())
            }
            __p((seen).toString())
            __p((millis >= 0).toString())
        
__check("500500\ntrue\ntrue")
}
