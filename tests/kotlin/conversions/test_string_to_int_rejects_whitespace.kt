// vybe-test: kotlin/conversions/test_string_to_int_rejects_whitespace
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun safe(value: String): Int {
            return try {
                value.toInt()
            } catch (e: Exception) {
                -1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((safe(" 12")).toString(), "-1")
            __check((safe("12 ")).toString(), "-1")
            __check((safe("\t12")).toString(), "-1")
        }
