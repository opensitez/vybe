// vybe-test: kotlin/type_casts/test_safe_cast_of_nullables
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun valueOrDefault(value: Any?): Int { val text = value as? Int
return text ?: 11 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((valueOrDefault(null)).toString(), "11")
__check((valueOrDefault(8)).toString(), "8") }
