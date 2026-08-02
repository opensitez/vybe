// vybe-test: kotlin/enums/test_enum_value_equality_and_comparison
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Toggle { OFF, ON }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Toggle.ON
            val b = Toggle.ON
            if (a == b) {
                __check(("same").toString(), "same")
            }
        }
