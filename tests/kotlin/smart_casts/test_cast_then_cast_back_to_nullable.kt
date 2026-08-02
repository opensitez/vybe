// vybe-test: kotlin/smart_casts/test_cast_then_cast_back_to_nullable
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "hello"
            val direct: String? = value as String
            val again: String? = direct as? String
            __check((direct == again).toString(), "true")
            __check((again?.length).toString(), "5")
        }
