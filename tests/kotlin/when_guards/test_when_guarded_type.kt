// vybe-test: kotlin/when_guards/test_when_guarded_type
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun describe(x: Any): String = when {
            x is Int && x > 10 -> "large-int"
            x is Int -> "int"
            x is String && x.isEmpty() -> "empty-str"
            x is String -> "string"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(11)).toString(), "large-int")
            __check((describe(3)).toString(), "int")
            __check((describe("")).toString(), "empty-str")
            __check((describe("x")).toString(), "string")
        }
