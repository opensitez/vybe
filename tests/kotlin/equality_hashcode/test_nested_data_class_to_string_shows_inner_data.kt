// vybe-test: kotlin/equality_hashcode/test_nested_data_class_to_string_shows_inner_data
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Inner(val value: String)
        data class Outer(val inner: Inner, val active: Boolean)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Outer(Inner("x"), true).toString()).toString(), "Outer(inner=Inner(value=x), active=true)")
        }
