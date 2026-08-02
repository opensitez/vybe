// vybe-test: kotlin/enums/test_enum_all_entries
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Level { LOW, MEDIUM, HIGH }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Level.LOW).toString(), "0")
            __check((Level.MEDIUM).toString(), "1")
            __check((Level.HIGH).toString(), "2")
        }
