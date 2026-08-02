// vybe-test: kotlin/scope/test_scope_after_smart_cast_branch_isolated_by_type
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun label(value: Any): String {
            return when (value) {
                is String -> "str:" + value.length
                is Int -> "int:" + value
                is Boolean -> "bool:" + value
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
            __check((label("abc")).toString(), "str:3")
            __check((label(9)).toString(), "int:9")
            __check((label(true)).toString(), "bool:true")
            __check((label(2.5)).toString(), "other")
        }
