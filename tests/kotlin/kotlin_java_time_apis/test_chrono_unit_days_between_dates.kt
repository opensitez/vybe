// vybe-test: kotlin/kotlin_java_time_apis/test_chrono_unit_days_between_dates
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

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
            val a = java.time.LocalDate.parse("2024-01-01")
            val b = java.time.LocalDate.parse("2024-01-10")
            __p((java.time.temporal.ChronoUnit.DAYS.between(a, b)).toString())
            __p((java.time.temporal.ChronoUnit.MONTHS.between(a, b)).toString())
            __p((java.time.temporal.ChronoUnit.WEEKS.between(a, b)).toString())
        
__check("9\n0\n1")
}
