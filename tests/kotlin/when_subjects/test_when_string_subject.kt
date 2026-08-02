// vybe-test: kotlin/when_subjects/test_when_string_subject
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun describe(input: String): String = when (input.length) {
            0 -> "empty"
            in 1..3 -> "short"
            else -> "long"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("")).toString(), "empty")
            __check((describe("ok")).toString(), "short")
            __check((describe("hello")).toString(), "long")
        }
