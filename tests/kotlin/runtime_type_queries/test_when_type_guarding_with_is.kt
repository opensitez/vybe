// vybe-test: kotlin/runtime_type_queries/test_when_type_guarding_with_is
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<Any> = listOf("k", 12, 3.4)
            val tags = values.map { value ->
                when (value) {
                    is Int -> "int"
                    is Double -> "double"
                    is String -> "string"
                    else -> "other"
                }
            }
            __check((tags.joinToString(",")).toString(), "string,int,double")
        }
