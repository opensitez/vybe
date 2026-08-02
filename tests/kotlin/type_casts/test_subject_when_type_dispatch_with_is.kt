// vybe-test: kotlin/type_casts/test_subject_when_type_dispatch_with_is
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun describe(value: Any): String {
            return when (value) {
                is Int -> if (value % 2 == 0) "even" else "odd"
                is String -> "len:" + value.length
                is Boolean -> "bool:" + (if (value) 1 else 0)
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
            __check((describe(4)).toString(), "even")
            __check((describe("go")).toString(), "len:2")
            __check((describe(false)).toString(), "bool:0")
            __check((describe(3.14)).toString(), "other")
        }
