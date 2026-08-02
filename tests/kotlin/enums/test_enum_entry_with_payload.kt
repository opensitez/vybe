// vybe-test: kotlin/enums/test_enum_entry_with_payload
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Planet(val order: Int) {
            MERCURY(1),
            VENUS(2),
            EARTH(3)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Planet.VENUS.order).toString(), "2")
            __check((Planet.EARTH.order + Planet.MERCURY.order).toString(), "4")
        }
