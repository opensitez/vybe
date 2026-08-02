// vybe-test: kotlin/infix/test_infix_custom_boolean
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Guard {
            infix fun allows(hour: Int): Boolean {
                return hour >= 9 && hour <= 17
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val shift = Guard()
            __check((shift.allows(10)).toString(), "true")
            __check((shift.allows(2)).toString(), "false")
        }
