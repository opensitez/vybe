// vybe-test: kotlin/type_casts/test_safe_cast_array_to_incompatible_component_type_is_null
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = arrayOf("a", "b")
            val casted = value as? Array<Int>
            __check((casted == null).toString(), "true")
        }
