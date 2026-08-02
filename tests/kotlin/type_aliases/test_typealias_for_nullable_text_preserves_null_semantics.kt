// vybe-test: kotlin/type_aliases/test_typealias_for_nullable_text_preserves_null_semantics
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Text = String?

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Text = null
            __check((value == null).toString(), "true")
            __check(((value ?: "fallback")).toString(), "fallback")
        }
