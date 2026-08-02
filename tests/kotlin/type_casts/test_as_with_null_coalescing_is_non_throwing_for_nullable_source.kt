// vybe-test: kotlin/type_casts/test_as_with_null_coalescing_is_non_throwing_for_nullable_source
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun extract(value: Any?): String {
            return (value as? String) ?: "missing"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((extract("kotlin")).toString(), "kotlin")
            __check((extract(null)).toString(), "missing")
        }
