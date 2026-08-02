// vybe-test: kotlin/printing/test_printing_uses_nullability_projection_in_template
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val maybe: String? = null
            __check(("value=${maybe ?: \"missing\"}").toString(), "value=missing")
            val present: String? = "ok"
            __check(("value=${present ?: \"missing\"}").toString(), "value=ok")
        }
