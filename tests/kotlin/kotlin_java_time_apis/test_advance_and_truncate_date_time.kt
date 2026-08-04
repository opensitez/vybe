// vybe-test: kotlin/kotlin_java_time_apis/test_advance_and_truncate_date_time
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
            val value = java.time.LocalDateTime.parse("2024-01-01T10:59:59")
            __p((value.plusSeconds(62).toString()).toString())
            __p((value.with(java.time.temporal.ChronoField.HOUR_OF_DAY, 0).toString()).toString())
        
__check("2024-01-01T11:01:01\n2024-01-01T00:59:59")
}
