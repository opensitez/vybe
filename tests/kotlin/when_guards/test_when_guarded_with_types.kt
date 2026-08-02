// vybe-test: kotlin/when_guards/test_when_guarded_with_types
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(v: Any): String = when {
            v is String && v.isEmpty() -> "empty"
            v is String -> "str"
            v is Int && v > 10 -> "big-int"
            v is Int -> "int"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label("")).toString(), "empty")
            __check((label("x")).toString(), "str")
            __check((label(11)).toString(), "big-int")
            __check((label(5)).toString(), "int")
        }
