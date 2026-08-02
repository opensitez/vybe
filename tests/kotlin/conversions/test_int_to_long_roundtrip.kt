// vybe-test: kotlin/conversions/test_int_to_long_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Int = 42
            val widened = source.toLong()
            val narrowed = widened.toInt()
            __check((widened).toString(), "42")
            __check((narrowed).toString(), "42")
            __check((narrowed == source).toString(), "true")
        }
