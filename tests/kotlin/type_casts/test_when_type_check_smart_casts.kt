// vybe-test: kotlin/type_casts/test_when_type_check_smart_casts
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun describe(value: Any): String {
            return when {
                value is String -> "string:" + value.length
                value is Int -> "int:" + (value + 1)
                value is Boolean -> "bool:" + (if (value) 1 else 0)
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("kotlin")).toString(), "string:6")
            __check((describe(6)).toString(), "int:7")
            __check((describe(false)).toString(), "bool:0")
            __check((describe(1.5)).toString(), "other")
        }
