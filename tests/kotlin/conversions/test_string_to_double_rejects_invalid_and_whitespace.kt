// vybe-test: kotlin/conversions/test_string_to_double_rejects_invalid_and_whitespace
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun safe(value: String): Double {
            return try {
                value.toDouble()
            } catch (e: Exception) {
                Double.NaN
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((safe("nan")).toString(), "NaN")
            __check((safe(" 2.0")).toString(), "NaN")
            __check((safe("bad")).toString(), "NaN")
        }
