// vybe-test: kotlin/enums/test_enum_reference_in_boolean
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Flag { TRUE, FALSE }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = Flag.TRUE
            __check((f == Flag.TRUE).toString(), "true")
            __check((f != Flag.FALSE).toString(), "true")
        }
