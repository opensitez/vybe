// vybe-test: kotlin/smart_casts/test_as_question_mark_on_incompatible_numeric_type_chain
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 9L
            val asInt: Int? = value as? Int
            val asLong: Long? = value as? Long
            __check((asInt == null).toString(), "true")
            __check((asLong != null).toString(), "true")
            __check((asLong?.toString() ?: "none").toString(), "9")
        }
