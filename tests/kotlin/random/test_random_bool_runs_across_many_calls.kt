// vybe-test: kotlin/random/test_random_bool_runs_across_many_calls
// origin: languages/kotlin/tests/kotlin/test_random.rs

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
            val r1 = kotlin.random.Random(107)
            val r2 = kotlin.random.Random(107)
            var trueCount = 0
            var i = 0
            while (i < 8) {
                if (r1.nextBoolean()) trueCount++
                i++
            }
            var trueCount2 = 0
            var j = 0
            while (j < 8) {
                if (r2.nextBoolean()) trueCount2++
                j++
            }
            __p((trueCount == trueCount2).toString())
            __p((trueCount in 0..8).toString())
        
__check("true\ntrue")
}
