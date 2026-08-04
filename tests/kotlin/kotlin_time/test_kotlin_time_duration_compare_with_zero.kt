// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_compare_with_zero
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

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
            val positive = 1.toDuration(DurationUnit.SECONDS)
            val zero = Duration.ZERO
            val negative = -(500.toDuration(DurationUnit.MILLISECONDS))
            __p((positive > zero).toString())
            __p((zero > negative).toString())
            __p((negative < zero).toString())
        
__check("true\ntrue\ntrue")
}
