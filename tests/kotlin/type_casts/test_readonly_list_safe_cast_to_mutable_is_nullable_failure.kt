// vybe-test: kotlin/type_casts/test_readonly_list_safe_cast_to_mutable_is_nullable_failure
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val readonly: Any = listOf("x", "y")
            val casted = readonly as? MutableList<String>
            __check((casted == null).toString(), "true")
        }
