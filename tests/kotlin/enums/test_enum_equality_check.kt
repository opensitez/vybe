// vybe-test: kotlin/enums/test_enum_equality_check
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class State { OFF, ON }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s1 = State.OFF
            val s2 = State.ON
            if (s1 != s2) {
                __check(("different states").toString(), "different states")
            }
        }
