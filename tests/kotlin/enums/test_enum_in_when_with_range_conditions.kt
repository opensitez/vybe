// vybe-test: kotlin/enums/test_enum_in_when_with_range_conditions
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Grade { A, B, C, D }

        fun label(g: Grade): String {
            return when (g) {
                Grade.A -> "excellent"
                Grade.B -> "good"
                Grade.C -> "ok"
                Grade.D -> "need work"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(Grade.A)).toString(), "excellent")
            __check((label(Grade.D)).toString(), "need work")
        }
