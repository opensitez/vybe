// vybe-test: kotlin/enums/test_enum_payloads_and_references
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Planet(val mass: Int) {
            MERCURY(1),
            EARTH(2),
            MARS(3)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val current = Planet.EARTH
            __check((current.mass).toString(), "2")
            __check((Planet.MARS.mass).toString(), "3")
        }
