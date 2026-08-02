// vybe-test: kotlin/when_guards/test_when_subject_with_let
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun classify(x: Int?): String = x?.let {
            when {
                it < 0 -> "neg"
                it == 0 -> "zero"
                else -> "pos"
            }
        } ?: "null"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(-2)).toString(), "neg")
            __check((classify(null)).toString(), "null")
        }
