// vybe-test: kotlin/enums/test_enum_properties_and_accessors
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Level(val code: Int) {
            LOW(1),
            MEDIUM(2),
            HIGH(3)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Level.MEDIUM.code).toString(), "2")
            __check((Level.HIGH.code).toString(), "3")
        }
