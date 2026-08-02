// vybe-test: kotlin/enums/test_enum_by_values_index_matches_ordinal
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Grade { A, B, C, D }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Grade.values()[0] == Grade.A).toString(), "true")
            __check((Grade.values()[2].ordinal).toString(), "2")
            __check((Grade.values().size).toString(), "4")
        }
