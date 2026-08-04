// vybe-test: kotlin/kotlin_java_time_apis/test_zoned_date_time_with_fixed_offset
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
            val zone = java.time.ZoneId.of("UTC")
            val value = java.time.ZonedDateTime.of(
                java.time.LocalDateTime.of(2024, 1, 1, 12, 0),
                zone
            )
            __p((value.toOffsetDateTime().offset.id).toString())
            __p((value.toInstant().toEpochMilli()).toString())
        
__check("Z\n1704110400000")
}
