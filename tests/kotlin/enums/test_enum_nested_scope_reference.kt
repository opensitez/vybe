// vybe-test: kotlin/enums/test_enum_nested_scope_reference
// origin: languages/kotlin/tests/kotlin/test_enums.rs

class Traffic {
            enum class Light { RED, YELLOW, GREEN }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val current = Traffic.Light.GREEN
            __check((current.name).toString(), "GREEN")
            __check((current.ordinal).toString(), "2")
        }
