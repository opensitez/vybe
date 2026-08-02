// vybe-test: kotlin/type_casts/test_cast_with_optional_source
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun toNumber(value: Any?): Int { return (value as? Int) ?: 0 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((toNumber(null)).toString(), "0")
__check((toNumber(4)).toString(), "4") }
