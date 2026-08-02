// vybe-test: kotlin/type_casts/test_as_cast_on_any_reference
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val any: Any = 77
            val num = any as Int
            __check((num + 1).toString(), "78")
        }
