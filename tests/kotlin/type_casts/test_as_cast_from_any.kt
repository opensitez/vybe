// vybe-test: kotlin/type_casts/test_as_cast_from_any
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val value: Any = 9
val casted = value as Int
__check((casted + 10).toString(), "19") }
