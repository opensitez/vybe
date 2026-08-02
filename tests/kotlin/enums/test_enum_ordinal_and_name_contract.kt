// vybe-test: kotlin/enums/test_enum_ordinal_and_name_contract
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Planet { MERCURY, EARTH, MARS }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val target = Planet.EARTH
            __check((target.name).toString(), "EARTH")
            __check((target.ordinal).toString(), "1")
            __check((Planet.MERCURY.ordinal < Planet.EARTH.ordinal).toString(), "true")
        }
