// vybe-test: kotlin/kotlin_time/test_kotlin_time_non_trivial_unit_rounding_ms_to_s
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1999.toDuration(DurationUnit.MILLISECONDS)
            __check((value.toLong(DurationUnit.SECONDS)).toString(), "1")
            __check((value.toDouble(DurationUnit.SECONDS)).toString(), "1.999")
        }
