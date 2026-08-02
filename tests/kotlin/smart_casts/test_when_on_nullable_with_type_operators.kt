// vybe-test: kotlin/smart_casts/test_when_on_nullable_with_type_operators
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<Any?> = listOf(null, "a", 3)
            val labels = values.map { item ->
                when (item) {
                    null -> "null"
                    is String -> "str"
                    is Int -> "int"
                    else -> "other"
                }
            }
            __check((labels.joinToString(",")).toString(), "null,str,int")
        }
