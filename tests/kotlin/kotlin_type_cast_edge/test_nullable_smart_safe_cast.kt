// vybe-test: kotlin/kotlin_type_cast_edge/test_nullable_smart_safe_cast
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_cast_edge.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val maybe: Any? = null
            __check(((maybe as? String)).toString(), "null")
            __check((("ok" as? String)).toString(), "ok")
        }
