// vybe-test: kotlin/literals/test_null_literal_and_type_inference
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            val safe = value ?: "none"
            __check((value == null).toString(), "true")
            __check((safe).toString(), "none")
            val explicit: String = "ok"
            __check((explicit).toString(), "ok")
        }
