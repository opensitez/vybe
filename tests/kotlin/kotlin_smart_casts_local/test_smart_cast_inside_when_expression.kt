// vybe-test: kotlin/kotlin_smart_casts_local/test_smart_cast_inside_when_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_smart_casts_local.rs

fun score(value: Any): String = when (value) {
            is String -> "s:" + value.length
            is Double -> "d:" + value.toInt()
            is Boolean -> "b:" + value
            else -> "n"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score("abc")).toString(), "s:3")
            __check((score(4.9)).toString(), "d:4")
            __check((score(false)).toString(), "b:false")
        }
